# Agente de Relatórios

Aplicação Streamlit que recebe uma planilha XLSX e gera:

- `RELATORIO_FINAL_<NOME_DO_ARQUIVO>.docx`: relatório Word com capa, logo da Usinox, sumário e uma página por registro;
- `RELATORIO_DESCIDAS_<NOME_DO_ARQUIVO>.xlsx`: relatório de legendas com as abas `TRANSPORTADORES`, `COMANDO` e `RAMAL DE ALIMENTAÇÃO`;
- um ZIP contendo os dois resultados.

## Recursos da versão atual

- o nome do projeto é extraído automaticamente do nome do arquivo enviado;
- a capa do Word usa o logo da Usinox;
- o rodapé do Word usa o logo da Usinox e paginação automática;
- o Word inclui um sumário com a página estimada de cada acionamento;
- a coluna artificial `DESCIDA` foi removida do Excel.

## Segurança

A aplicação executa apenas a rotina fixa existente em `processor.py`. Ela não utiliza `exec()` e não executa código enviado por usuários.

## Executar no computador

```bash
python -m venv .venv
```

Ative o ambiente virtual e instale as dependências:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar gratuitamente no Streamlit Community Cloud

1. Crie um repositório no GitHub.
2. Envie todos os arquivos desta pasta ao repositório.
3. Acesse o Streamlit Community Cloud e conecte sua conta do GitHub.
4. Selecione o repositório, a branch principal e o arquivo `app.py`.
5. Clique em **Deploy**.

A aplicação receberá um endereço `streamlit.app` que poderá ser compartilhado.

## Estrutura

```text
agente_relatorios/
├── app.py
├── processor.py
├── logo_usinox.jpg
├── requirements.txt
├── supabase_schema.sql
├── .gitignore
├── .streamlit/
│   ├── config.toml
│   └── secrets.toml.example
└── README.md
```

## Supabase

O arquivo `supabase_schema.sql` contém uma tabela inicial para histórico de execuções. A integração foi deixada desativada no MVP para evitar armazenamento de dados antes de configurar autenticação e políticas de segurança.

## Regras preservadas do código original

- a primeira aba do arquivo enviado é usada como base;
- linhas totalmente vazias são removidas;
- o Word mantém uma página por registro de acionamento;
- registros com `POS` contendo `WF` vão para `RAMAL DE ALIMENTAÇÃO`;
- registros com `WM` preenchido, exceto `WF`, vão para `TRANSPORTADORES`;
- registros de `WC`, `WN`, `WS` ou com campos de comando preenchidos vão para `COMANDO`;
- a coluna artificial `DESCIDA` foi removida; cada aba usa o identificador real do equipamento (`WM`, `WC`/`WN`/`WS` ou `WF`).
