# BI de Falta de Bipagem (Streamlit)

## Rodar no seu computador
1. Instale o Python 3.10+.
2. Na pasta do projeto: `pip install -r requirements.txt`
3. Execute: `streamlit run app.py` (abre em http://localhost:8501)
4. Na barra lateral, envie o(s) relatório(s): CSV, TSV, XLS, XLSX, XLSM ou ODS.

## Publicar no Streamlit Community Cloud (link compartilhável)
1. Crie um repositório no GitHub com `app.py` e `requirements.txt` (de preferência **privado**).
2. Acesse https://share.streamlit.io, entre com o GitHub e clique em **Create app**.
3. Escolha o repositório, branch `main` e o arquivo principal `app.py`. Clique em **Deploy**.
4. Compartilhe o link. Em apps privados, convide os e-mails de quem pode acessar.

> Os relatórios têm nome de clientes e operadores: não deixe o app público.

## O que o app espera
- Aba/arquivo de **detalhe**: precisa ter a coluna `Número da Remessa` e `Horário do bipe de descarga do veículo recebido`.
- Aba **Resumo** (opcional): colunas `Código SC`, `Nome SC`, `Qtd pedidos a bipar`,
  `Qtd pedidos não bipados no recebimento`, `Qtd pedidos não bipados na expedição`. Sem ela, o app mostra só as contagens.
- Se os nomes das colunas mudarem, ajuste as constantes no bloco 2 do `app.py`.
