# Finanças da Casa

Automação que lê a fatura do cartão em PDF (salva numa pasta do Google
Drive), categoriza os lançamentos e grava os totais na planilha
"Finanças da Família 2026" (Google Sheets).

## Como usar

1. Instale as dependências:

   ```bash
   pip install -r requirements.txt
   ```

2. Configure uma Service Account no [Google Cloud Console](https://console.cloud.google.com):
   - Crie um projeto (ou use um existente)
   - Ative a **Google Sheets API** e a **Google Drive API**
   - Crie uma Service Account, gere uma chave JSON e salve como
     `credentials.json` na raiz do projeto (esse arquivo **nunca** deve ir
     para o GitHub em texto puro — o `.gitignore` já o exclui; use GitHub
     Secrets para automações na nuvem)
   - Compartilhe a planilha "Finanças da Família 2026" e a pasta de
     faturas no Drive com o e-mail da service account
     (algo como `xxxx@yyyy.iam.gserviceaccount.com`), dando permissão de
     Editor nos dois

3. Ajuste as constantes no topo de `atualizar_planilha_financas.py`
   (`SPREADSHEET_ID`, `DRIVE_FOLDER_ID`, `SHEET_NAME`, `COLUNA_DO_MES`,
   `LINHA_DO_ITEM`) para bater com a estrutura real da planilha.

4. Teste manualmente com um PDF local antes de automatizar:

   ```bash
   python atualizar_planilha_financas.py --pdf "fatura_teste.pdf" --mes agosto
   ```

5. Teste o modo automático (lê a pasta do Drive) em modo simulação:

   ```bash
   python atualizar_planilha_financas.py --mes agosto
   ```

6. Quando validado, use `--escrever` para gravar de verdade na planilha e
   marcar os PDFs como processados no Drive:

   ```bash
   python atualizar_planilha_financas.py --mes agosto --escrever
   ```

7. Para rodar sozinho, sem depender de ninguém passar `--mes` na mão, crie
   uma Routine (claude.ai/code/routines ou `/schedule` no Claude Code CLI)
   apontando para este repositório, na frequência desejada (semanal / a
   cada 10 dias). Quando `--mes` é omitido o script usa automaticamente o
   mês corrente (`mes_atual()`), então a Routine pode chamar apenas:

   ```bash
   python atualizar_planilha_financas.py --escrever
   ```

## Aviso

A extração de PDF e as regras de categorização em `REGRAS` são um ponto de
partida baseado nos extratos do Ourocard Platinum Estilo (BB). Ajuste as
palavras-chave conforme novos comerciantes forem aparecendo nas faturas.
