from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import math
from pathlib import Path
from typing import Any

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


SUMMARY_ITEMS_PER_PAGE = 42
LOGO_FILENAME = "logo_usinox.jpg"


@dataclass(frozen=True)
class ResultadoProcessamento:
    docx: bytes
    xlsx: bytes
    total_registros: int
    total_transportadores: int
    total_comando: int
    total_ramal: int
    avisos: list[str]


def _is_nan(value: Any) -> bool:
    try:
        return bool(pd.isna(value)) or (
            isinstance(value, float) and math.isnan(value)
        )
    except (TypeError, ValueError):
        return False


def _is_blank(value: Any) -> bool:
    if _is_nan(value):
        return True
    return isinstance(value, str) and value.strip() == ""


def _safe_str(value: Any) -> str:
    if _is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    normalized = {
        str(column).strip().upper(): column
        for column in df.columns
    }
    for candidate in candidates:
        found = normalized.get(candidate.strip().upper())
        if found is not None:
            return found
    return None


def _get_val(row: pd.Series, *columns: str | None) -> Any:
    for column in columns:
        if column is not None and column in row.index:
            value = row.get(column, "")
            return "" if _is_nan(value) else value
    return ""


def _set_cell_background(cell, color_hex: str) -> None:
    cell_properties = cell._tc.get_or_add_tcPr()
    shading = cell_properties.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        cell_properties.append(shading)
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), color_hex)


def _set_cell_text(
    cell,
    text: Any,
    *,
    bold: bool = False,
    size_pt: int = 10,
) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(_safe_str(text))
    run.bold = bold
    run.font.size = Pt(size_pt)
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")


def _set_col_widths(table, widths_cm: list[float]) -> None:
    table.autofit = False
    for row in table.rows:
        for index, width in enumerate(widths_cm):
            row.cells[index].width = Cm(width)


def _add_title_borders(
    paragraph,
    *,
    color_hex: str = "F28C28",
    size_eighth_points: int = 12,
    left_space: str = "4",
) -> None:
    paragraph_properties = paragraph._p.get_or_add_pPr()
    borders = paragraph_properties.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        paragraph_properties.append(borders)

    for tag in ("top", "right", "bottom", "left", "between"):
        existing = borders.find(qn(f"w:{tag}"))
        if existing is not None:
            borders.remove(existing)

    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size_eighth_points))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color_hex)
    borders.append(bottom)

    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size_eighth_points))
    left.set(qn("w:space"), left_space)
    left.set(qn("w:color"), color_hex)
    borders.append(left)


def _set_table_dashed_borders(
    table,
    *,
    size_eighth_points: int = 2,
    color: str = "auto",
) -> None:
    table_properties = table._tbl.tblPr
    borders = table_properties.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        table_properties.append(borders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "dashed")
        element.set(qn("w:sz"), str(size_eighth_points))
        element.set(qn("w:color"), color)
        element.set(qn("w:space"), "0")


def _add_field(run, field_code: str) -> None:
    run_xml = run._r
    field_begin = OxmlElement("w:fldChar")
    field_begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = field_code
    field_separator = OxmlElement("w:fldChar")
    field_separator.set(qn("w:fldCharType"), "separate")
    field_end = OxmlElement("w:fldChar")
    field_end.set(qn("w:fldCharType"), "end")
    run_xml.extend([field_begin, instruction, field_separator, field_end])


def _prepare_dataframe(excel_bytes: bytes) -> pd.DataFrame:
    try:
        dataframe = pd.read_excel(BytesIO(excel_bytes), sheet_name=0)
    except Exception as exc:
        raise ValueError(
            "Não foi possível ler o arquivo Excel. Confirme se ele é um .xlsx válido."
        ) from exc

    blank_mask = dataframe.apply(lambda column: column.map(_is_blank))
    dataframe = dataframe.loc[~blank_mask.all(axis=1)].reset_index(drop=True)

    if dataframe.empty:
        raise ValueError("A primeira aba da planilha não contém registros válidos.")

    return dataframe


def _build_titles(
    dataframe: pd.DataFrame,
    position_column: str | None,
    description_column: str | None,
) -> list[str]:
    titles: list[str] = []
    for _, row in dataframe.iterrows():
        position = _safe_str(row.get(position_column, "")) if position_column else ""
        description = (
            _safe_str(row.get(description_column, ""))
            if description_column
            else ""
        )

        if position and description:
            titles.append(f"{position} - {description}")
        elif description:
            titles.append(description)
        elif position:
            titles.append(position)
        else:
            titles.append("REGISTRO")
    return titles


def _configure_section(section) -> None:
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.header_distance = Cm(0.7)
    section.footer_distance = Cm(0.7)


def _clear_cell(cell) -> None:
    cell.text = ""
    if cell.paragraphs:
        paragraph = cell.paragraphs[0]
        for run in list(paragraph.runs):
            run._element.getparent().remove(run._element)


def _insert_logo(paragraph, image_path: Path, width_cm: float) -> None:
    if image_path.exists():
        run = paragraph.add_run()
        run.add_picture(str(image_path), width=Cm(width_cm))


def _setup_header_footer(section, project_name: str, logo_path: Path) -> None:
    header = section.header
    header_paragraph = header.paragraphs[0]
    header_paragraph.text = ""
    header_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = header_paragraph.add_run("RELATÓRIO DE ACIONAMENTOS")
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.name = "Calibri"

    subtitle_paragraph = header.add_paragraph()
    subtitle_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    subtitle_run = subtitle_paragraph.add_run(project_name)
    subtitle_run.font.size = Pt(8)
    subtitle_run.font.name = "Calibri"
    subtitle_run.font.color.rgb = RGBColor(80, 80, 80)

    footer = section.footer
    table_width = section.page_width - section.left_margin - section.right_margin
    footer_table = footer.add_table(rows=1, cols=2, width=table_width)
    footer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    footer_table.autofit = False

    left_cell = footer_table.cell(0, 0)
    right_cell = footer_table.cell(0, 1)
    left_cell.width = Cm(4)
    right_cell.width = Cm(12)

    _clear_cell(left_cell)
    _clear_cell(right_cell)

    left_paragraph = left_cell.paragraphs[0]
    left_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _insert_logo(left_paragraph, logo_path, width_cm=2.2)

    right_paragraph = right_cell.paragraphs[0]
    right_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    right_paragraph.add_run("Página ")
    _add_field(right_paragraph.add_run(), "PAGE  \\* MERGEFORMAT")
    right_paragraph.add_run(" de ")
    _add_field(right_paragraph.add_run(), "NUMPAGES  \\* MERGEFORMAT")
    for run in right_paragraph.runs:
        run.font.size = Pt(9)
        run.font.name = "Calibri"


def _build_docx(dataframe: pd.DataFrame, project_name: str) -> bytes:
    dataframe_title = dataframe.copy()
    dataframe_table = dataframe.copy()

    columns_to_remove = {
        "DESCRICAO",
        "DESCRIÇÃO",
        "Q",
        "ID",
        "IP (TD1)",
        "MASK (TD2)",
        "POS",
    }
    for column in list(dataframe_table.columns):
        if str(column).strip().upper() in columns_to_remove:
            dataframe_table = dataframe_table.drop(columns=[column])

    limit_index: int | None = None
    for index, column in enumerate(dataframe_table.columns):
        if str(column).strip().upper() == "OBS_AUTOMACAO":
            limit_index = index
            break

    labels = (
        dataframe_table.columns[: limit_index + 1].tolist()
        if limit_index is not None
        else dataframe_table.columns.tolist()
    )

    for extra in ("DESTINO", "ORIGEM", "WS"):
        if extra not in labels:
            labels.append(extra)

    position_column = _find_col(dataframe_title, ["POS"])
    description_column = _find_col(
        dataframe_title,
        ["DESCRICAO", "DESCRIÇÃO"],
    )
    titles = _build_titles(dataframe_title, position_column, description_column)

    document = Document()
    section = document.sections[0]
    _configure_section(section)
    logo_path = Path(__file__).resolve().parent / LOGO_FILENAME
    _setup_header_footer(section, project_name, logo_path)

    # Capa
    cover_title = document.add_paragraph()
    cover_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cover_title.paragraph_format.space_before = Pt(20)
    _insert_logo(cover_title, logo_path, width_cm=12.0)

    for _ in range(3):
        document.add_paragraph()

    title_paragraph = document.add_paragraph()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_paragraph.add_run("RELATÓRIO DE ACIONAMENTOS")
    title_run.bold = True
    title_run.font.size = Pt(22)
    title_run.font.name = "Calibri"

    project_paragraph = document.add_paragraph()
    project_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    project_run = project_paragraph.add_run(project_name)
    project_run.bold = True
    project_run.font.size = Pt(18)
    project_run.font.name = "Calibri"
    project_run.font.color.rgb = RGBColor(0, 91, 170)

    info_paragraph = document.add_paragraph()
    info_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_run = info_paragraph.add_run("Documento gerado automaticamente pelo Agente de Relatórios")
    info_run.font.size = Pt(10)
    info_run.font.name = "Calibri"
    info_run.font.color.rgb = RGBColor(90, 90, 90)

    document.add_page_break()

    # Sumário manual
    summary_page_count = max(1, math.ceil(len(titles) / SUMMARY_ITEMS_PER_PAGE))
    records_start_page = summary_page_count + 2

    for page_idx in range(summary_page_count):
        if page_idx > 0:
            document.add_page_break()

        heading = document.add_heading(
            "SUMÁRIO" if page_idx == 0 else "SUMÁRIO (continuação)",
            level=1,
        )
        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in heading.runs:
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.name = "Calibri"
        _add_title_borders(heading)

        subtitle = document.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.LEFT
        subtitle_run = subtitle.add_run(f"Projeto: {project_name}")
        subtitle_run.font.size = Pt(10)
        subtitle_run.font.name = "Calibri"

        start = page_idx * SUMMARY_ITEMS_PER_PAGE
        end = min(start + SUMMARY_ITEMS_PER_PAGE, len(titles))

        for local_idx, title in enumerate(titles[start:end]):
            global_idx = start + local_idx
            page_number = records_start_page + global_idx

            item = document.add_paragraph()
            item.paragraph_format.space_before = Pt(0)
            item.paragraph_format.space_after = Pt(0)
            item.paragraph_format.line_spacing = 1
            item.paragraph_format.tab_stops.add_tab_stop(
                Cm(16.0),
                WD_TAB_ALIGNMENT.RIGHT,
                WD_TAB_LEADER.DOTS,
            )

            title_run = item.add_run(title)
            title_run.font.size = Pt(8)
            title_run.font.name = "Calibri"

            page_run = item.add_run(f"\t{page_number}")
            page_run.font.size = Pt(8)
            page_run.font.name = "Calibri"

    if titles:
        document.add_page_break()

    # Páginas dos acionamentos
    for row_number in range(len(dataframe_table)):
        table_row = dataframe_table.iloc[row_number]
        title = titles[row_number]

        if row_number > 0:
            document.add_page_break()

        heading = document.add_heading(title, level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        heading.paragraph_format.space_before = Pt(0)
        heading.paragraph_format.space_after = Pt(6)
        for run in heading.runs:
            run.font.color.rgb = RGBColor(0, 0, 0)
            run.bold = True
            run.font.size = Pt(16)
            run.font.name = "Calibri"
        _add_title_borders(heading)

        table = document.add_table(rows=len(labels), cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        _set_col_widths(table, [4.0, 12.0])
        _set_table_dashed_borders(table)

        for index, label in enumerate(labels):
            value = table_row.get(label, "")
            if str(label).strip().upper() == "NS" and not _is_blank(value):
                try:
                    value = str(int(float(value)))
                except (TypeError, ValueError):
                    value = _safe_str(value)

            _set_cell_text(table.cell(index, 0), label, bold=True, size_pt=10)
            _set_cell_text(table.cell(index, 1), value, size_pt=10)
            background = "D9D9D9" if index % 2 == 0 else "FFFFFF"
            _set_cell_background(table.cell(index, 0), background)
            _set_cell_background(table.cell(index, 1), background)

    output = BytesIO()
    document.save(output)
    return output.getvalue()


def _build_xlsx(
    dataframe: pd.DataFrame,
) -> tuple[bytes, int, int, int]:
    base = dataframe.copy()

    column_wm = _find_col(base, ["WM"])
    column_wm_cable = _find_col(
        base,
        ["WM_CABO", "WM CABO", "WMCABO", "WM-CABO"],
    )
    column_location = _find_col(base, ["LOC"])
    column_description = _find_col(base, ["DESCRICAO", "DESCRIÇÃO"])
    column_origin = _find_col(base, ["ORIGEM"])
    column_destination = _find_col(base, ["DESTINO"])
    column_ws = _find_col(base, ["WS"])
    column_wc = _find_col(base, ["WC"])
    column_ws_cable = _find_col(
        base,
        ["WS_CABO", "WS CABO", "WSCABO", "WS-CABO"],
    )
    column_wc_cable = _find_col(
        base,
        ["WC_CABO", "WC CABO", "WCCABO", "WC-CABO"],
    )
    column_position = _find_col(base, ["POS"])

    workbook = Workbook()

    align_center = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )
    align_left = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=True,
    )
    font_bold_12 = Font(bold=True, size=12)
    font_bold_14 = Font(bold=True, size=14)
    fill_header = PatternFill("solid", fgColor="D9D9D9")
    fill_subtitle = PatternFill("solid", fgColor="F2F2F2")
    thin = Side(style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    def sheet_header(worksheet, number_of_columns: int, subtitle: str) -> None:
        worksheet.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=number_of_columns,
        )
        title_cell = worksheet.cell(
            row=1,
            column=1,
            value="RELATÓRIO DE LEGENDAS",
        )
        title_cell.font = font_bold_14
        title_cell.alignment = align_center

        worksheet.merge_cells(
            start_row=2,
            start_column=1,
            end_row=2,
            end_column=number_of_columns,
        )
        subtitle_cell = worksheet.cell(row=2, column=1, value=subtitle)
        subtitle_cell.font = font_bold_12
        subtitle_cell.alignment = align_center
        subtitle_cell.fill = fill_subtitle
        worksheet.row_dimensions[1].height = 22
        worksheet.row_dimensions[2].height = 18

    def write_header_row(worksheet, headers: list[str]) -> None:
        for column_number, header in enumerate(headers, start=1):
            cell = worksheet.cell(row=3, column=column_number, value=header)
            cell.font = font_bold_12
            cell.alignment = align_center
            cell.fill = fill_header
            cell.border = border_all
        worksheet.row_dimensions[3].height = 20

    def write_data_row(worksheet, row_number: int, values: list[Any]) -> None:
        for column_number, value in enumerate(values, start=1):
            cell = worksheet.cell(row=row_number, column=column_number, value=value)
            cell.alignment = align_center if column_number == 1 else align_left
            cell.border = border_all

    transport_sheet = workbook.active
    transport_sheet.title = "TRANSPORTADORES"
    transport_headers = [
        "WM",
        "WM_CABO",
        "LOC",
        "ORIGEM",
        "DESTINO",
        "DESCRIÇÃO",
    ]
    for column, width in zip("ABCDEF", [16, 18, 16, 18, 18, 60]):
        transport_sheet.column_dimensions[column].width = width
    sheet_header(transport_sheet, 6, "TRANSPORTADORES")
    write_header_row(transport_sheet, transport_headers)

    transport_count = 0
    output_row = 4
    for _, row in base.iterrows():
        wm_value = _get_val(row, column_wm)
        position_value = _get_val(row, column_position)
        contains_wf = (
            isinstance(position_value, str)
            and "WF" in position_value.upper()
        )
        if _is_blank(wm_value) or contains_wf:
            continue

        write_data_row(
            transport_sheet,
            output_row,
            [
                wm_value,
                _get_val(row, column_wm_cable),
                _get_val(row, column_location),
                _get_val(row, column_origin),
                _get_val(row, column_destination),
                _safe_str(_get_val(row, column_description)),
            ],
        )
        transport_count += 1
        output_row += 1

    transport_sheet.freeze_panes = "A4"
    transport_sheet.print_title_rows = "1:3"

    command_sheet = workbook.create_sheet(title="COMANDO")
    command_headers = [
        "COMANDO",
        "CABO",
        "LOC",
        "POS",
        "ORIGEM",
        "DESTINO",
        "DESCRIÇÃO",
    ]
    for column, width in zip("ABCDEFG", [16, 20, 16, 12, 18, 18, 60]):
        command_sheet.column_dimensions[column].width = width
    sheet_header(command_sheet, 7, "COMANDO")
    write_header_row(command_sheet, command_headers)

    command_count = 0
    output_row = 4
    for _, row in base.iterrows():
        ws_value = _get_val(row, column_ws)
        wc_value = _get_val(row, column_wc)
        position_value = _get_val(row, column_position)
        position_upper = (
            position_value.upper()
            if isinstance(position_value, str)
            else ""
        )

        enters_command = (
            "WC" in position_upper
            or "WN" in position_upper
            or not _is_blank(ws_value)
            or not _is_blank(wc_value)
        )
        if not enters_command:
            continue

        if not _is_blank(ws_value):
            principal = ws_value
            cable = _get_val(row, column_ws_cable)
        elif not _is_blank(wc_value):
            principal = wc_value
            cable = _get_val(row, column_wc_cable)
        elif not _is_blank(position_value):
            principal = position_value
            if "WS" in position_upper:
                cable = _get_val(row, column_ws_cable)
            elif "WC" in position_upper:
                cable = _get_val(row, column_wc_cable)
            else:
                cable = ""
        else:
            principal = ""
            cable = ""

        write_data_row(
            command_sheet,
            output_row,
            [
                principal,
                cable,
                _get_val(row, column_location),
                position_value,
                _get_val(row, column_origin),
                _get_val(row, column_destination),
                _safe_str(_get_val(row, column_description)),
            ],
        )
        command_count += 1
        output_row += 1

    if command_count:
        command_sheet.freeze_panes = "A4"
        command_sheet.print_title_rows = "1:3"
    else:
        command_sheet.sheet_state = "hidden"

    branch_sheet = workbook.create_sheet(title="RAMAL DE ALIMENTAÇÃO")
    branch_headers = [
        "ALIMENTAÇÃO (WF)",
        "CABO",
        "LOC",
        "ORIGEM",
        "DESTINO",
        "DESCRIÇÃO",
    ]
    for column, width in zip("ABCDEF", [20, 18, 16, 18, 18, 60]):
        branch_sheet.column_dimensions[column].width = width
    sheet_header(branch_sheet, 6, "RAMAL DE ALIMENTAÇÃO")
    write_header_row(branch_sheet, branch_headers)

    branch_count = 0
    output_row = 4
    for _, row in base.iterrows():
        position_value = _get_val(row, column_position)
        wm_value = _get_val(row, column_wm)
        contains_wf = (
            isinstance(position_value, str)
            and "WF" in position_value.upper()
        )
        if not (contains_wf and not _is_blank(wm_value)):
            continue

        write_data_row(
            branch_sheet,
            output_row,
            [
                position_value,
                _get_val(row, column_wm_cable),
                _get_val(row, column_location),
                _get_val(row, column_origin),
                _get_val(row, column_destination),
                _safe_str(_get_val(row, column_description)),
            ],
        )
        branch_count += 1
        output_row += 1

    if branch_count:
        branch_sheet.freeze_panes = "A4"
        branch_sheet.print_title_rows = "1:3"
    else:
        branch_sheet.sheet_state = "hidden"

    output = BytesIO()
    workbook.save(output)
    return output.getvalue(), transport_count, command_count, branch_count


def processar_documento(
    excel_bytes: bytes,
    project_name: str | None = None,
) -> ResultadoProcessamento:
    if not excel_bytes:
        raise ValueError("O arquivo enviado está vazio.")

    normalized_project_name = (project_name or "Projeto").strip() or "Projeto"

    dataframe = _prepare_dataframe(excel_bytes)
    document_bytes = _build_docx(dataframe, normalized_project_name)
    workbook_bytes, transport_count, command_count, branch_count = _build_xlsx(
        dataframe,
    )

    position_column = _find_col(dataframe, ["POS"])
    description_column = _find_col(dataframe, ["DESCRICAO", "DESCRIÇÃO"])
    warnings: list[str] = []

    if position_column is None:
        warnings.append("A coluna POS não foi encontrada.")
    if description_column is None:
        warnings.append("A coluna DESCRICAO/DESCRIÇÃO não foi encontrada.")

    if position_column is not None or description_column is not None:
        incomplete_count = 0
        for _, row in dataframe.iterrows():
            position = (
                _get_val(row, position_column)
                if position_column is not None
                else ""
            )
            description = (
                _get_val(row, description_column)
                if description_column is not None
                else ""
            )
            if _is_blank(position) and _is_blank(description):
                incomplete_count += 1
        if incomplete_count:
            warnings.append(
                f"{incomplete_count} registro(s) sem POS e sem descrição "
                "foram mantidos como 'REGISTRO'."
            )

    return ResultadoProcessamento(
        docx=document_bytes,
        xlsx=workbook_bytes,
        total_registros=len(dataframe),
        total_transportadores=transport_count,
        total_comando=command_count,
        total_ramal=branch_count,
        avisos=warnings,
    )
