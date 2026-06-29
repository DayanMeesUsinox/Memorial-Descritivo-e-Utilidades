from __future__ import annotations

from io import BytesIO
from pathlib import Path
import zipfile

import streamlit as st

from processor import processar_documento


MAX_FILE_SIZE_MB = 50


st.set_page_config(
    page_title="Agente de Relatórios",
    page_icon="📄",
    layout="centered",
)

st.title("Agente de Relatórios")
st.write(
    "Envie a planilha-base em formato XLSX. A aplicação gera o relatório "
    "de acionamentos em Word com capa, sumário e identidade visual da Usinox, "
    "além do relatório de legendas em Excel."
)

uploaded_file = st.file_uploader(
    "Planilha-base",
    type=["xlsx"],
    accept_multiple_files=False,
)

if uploaded_file is not None:
    file_size_mb = uploaded_file.size / (1024 * 1024)
    project_name = Path(uploaded_file.name).stem.strip() or "Projeto"
    safe_project_name = project_name.replace("/", "-").replace("\\", "-")
    st.caption(
        f"Arquivo: {uploaded_file.name} — {file_size_mb:.2f} MB"
        f" — Projeto: {project_name}"
    )

    if file_size_mb > MAX_FILE_SIZE_MB:
        st.error(
            f"O arquivo excede o limite de {MAX_FILE_SIZE_MB} MB definido "
            "para esta primeira versão."
        )
    elif st.button("Gerar relatórios", type="primary", use_container_width=True):
        with st.spinner("Processando a planilha..."):
            try:
                result = processar_documento(
                    uploaded_file.getvalue(),
                    project_name=project_name,
                )
            except ValueError as exc:
                st.error(str(exc))
            except Exception:
                st.exception(
                    RuntimeError(
                        "Ocorreu um erro inesperado durante o processamento."
                    )
                )
            else:
                st.success("Relatórios gerados com sucesso.")

                metric_1, metric_2, metric_3, metric_4 = st.columns(4)
                metric_1.metric("Registros", result.total_registros)
                metric_2.metric("Transportadores", result.total_transportadores)
                metric_3.metric("Comando", result.total_comando)
                metric_4.metric("Ramal", result.total_ramal)

                for warning in result.avisos:
                    st.warning(warning)

                docx_name = f"RELATORIO_FINAL_{safe_project_name}.docx"
                xlsx_name = f"RELATORIO_DESCIDAS_{safe_project_name}.xlsx"
                zip_name = f"RELATORIOS_GERADOS_{safe_project_name}.zip"

                zip_buffer = BytesIO()
                with zipfile.ZipFile(
                    zip_buffer,
                    mode="w",
                    compression=zipfile.ZIP_DEFLATED,
                ) as archive:
                    archive.writestr(docx_name, result.docx)
                    archive.writestr(xlsx_name, result.xlsx)

                st.download_button(
                    "Baixar relatório Word",
                    data=result.docx,
                    file_name=docx_name,
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    ),
                    use_container_width=True,
                )
                st.download_button(
                    "Baixar relatório Excel",
                    data=result.xlsx,
                    file_name=xlsx_name,
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )
                st.download_button(
                    "Baixar os dois arquivos em ZIP",
                    data=zip_buffer.getvalue(),
                    file_name=zip_name,
                    mime="application/zip",
                    use_container_width=True,
                )

st.divider()
st.caption(
    "Esta versão executa somente a rotina aprovada do projeto. Ela não "
    "executa código enviado pelos usuários."
)
