# Publicação do dashboard

## Streamlit Community Cloud

1. Crie um repositório no GitHub e envie os arquivos deste diretório.
2. Entre em `share.streamlit.io` e selecione **Create app**.
3. Escolha o repositório, a branch e informe `app.py` como arquivo principal.
4. Publique. A plataforma instalará as dependências de `requirements.txt`.

A planilha precisa permanecer compartilhada como “qualquer pessoa com o link — leitor”. A aplicação consulta o arquivo a cada dez minutos e usa `dados_projeto.xlsx` como contingência caso o Google esteja indisponível.

Para usar outra planilha, configure no Streamlit Cloud uma variável de ambiente chamada `GOOGLE_SHEET_ID` com o identificador do novo arquivo.

## Execução local

```powershell
python -m pip install -r requirements.txt
streamlit run app.py
```

## Google Apps Script

O Streamlit é a opção recomendada para este projeto porque o painel usa Python, pandas e Plotly. Apps Script não executa Python; para publicar por ele seria necessário reescrever o front-end em HTML/JavaScript e criar funções de leitura no Google Sheets.
