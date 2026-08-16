# Finanças da Casa

Automação que lê a fatura do cartão em PDF (salva numa pasta do Google
Drive), categoriza os lançamentos e grava os totais na planilha
"Finanças da Família 2026" (Google Sheets).

## Como usar

1. Instale as dependências:

   ```bash
   pip install -r requirements.txt
   ```

2. **Importante:** a planilha "Finanças da Família 2026" precisa estar em
   formato **nativo do Google Sheets**, não `.xlsx`. Se o arquivo foi
   criado por upload de um Excel, abra-o no Drive e use
   **Arquivo → Salvar como Planilhas Google** — isso cria uma cópia nova
   (com URL diferente) que a API do Sheets consegue manipular; arquivos
   `.xlsx` em "modo compatibilidade" não são acessíveis pela API.

3. Configure uma Service Account no [Google Cloud Console](https://console.cloud.google.com):
   - Crie um projeto (ou use um existente)
   - Ative a **Google Sheets API** e a **Google Drive API**
   - Crie uma Service Account, gere uma chave JSON e salve como
     `credentials.json` na raiz do projeto (esse arquivo **nunca** deve ir
     para o GitHub em texto puro — o `.gitignore` já o exclui; use GitHub
     Secrets para automações na nuvem)
   - Compartilhe a planilha "Finanças da Família 2026" (a versão nativa
     do passo 2) e a pasta de faturas no Drive com o e-mail da service
     account (algo como `xxxx@yyyy.iam.gserviceaccount.com`), dando
     permissão de Editor nos dois

   As constantes `SPREADSHEET_ID`, `DRIVE_FOLDER_ID` e `SHEET_NAME` no
   topo de `atualizar_planilha_financas.py` já estão preenchidas e
   confirmadas contra a estrutura real da planilha (aba `"2026"`, pasta
   `Faturas do Cartão da Casa`); só reajuste se a estrutura mudar.

4. (Opcional, se quiser testar localmente primeiro) Com o `credentials.json`
   na raiz do projeto, teste com um PDF local:

   ```bash
   python atualizar_planilha_financas.py --pdf "fatura_teste.pdf" --mes agosto
   ```

   Depois teste o modo automático (lê a pasta do Drive) em modo simulação:

   ```bash
   python atualizar_planilha_financas.py --mes agosto
   ```

   Quando validado, `--escrever` grava de verdade e marca os PDFs como
   processados no Drive:

   ```bash
   python atualizar_planilha_financas.py --mes agosto --escrever
   ```

5. **Automação via GitHub Actions** (roda sozinha, sem depender de
   nenhum computador ligado — já configurado em
   `.github/workflows/atualizar-planilha.yml`):

   - No GitHub, vá em **Settings → Secrets and variables → Actions →
     New repository secret**
   - Nome: `GOOGLE_CREDENTIALS_JSON`
   - Valor: cole o conteúdo inteiro do arquivo `credentials.json` baixado
     no passo 3
   - **Add secret**

   O workflow roda automaticamente toda segunda-feira às 09:00 (horário
   de Brasília) usando `--escrever`, e usa `mes_atual()` para saber qual
   mês gravar — não precisa passar `--mes` na mão. Ajuste o `cron` no
   arquivo do workflow se quiser outra frequência (ex.: conforme a data
   de fechamento da fatura).

   Para testar sem esperar a segunda-feira: aba **Actions** do
   repositório → **Atualizar planilha com fatura do cartão** →
   **Run workflow**. Deixe a opção "Gravar de verdade" desmarcada pra
   rodar em modo simulação primeiro.

6. **Painel do mês atual (mobile)** — uma página simples comparando o
   gasto do mês corrente com o orçamento (coluna "budget" da planilha),
   categoria por categoria, publicada automaticamente no GitHub Pages
   (`.github/workflows/atualizar-painel.yml`):

   - No GitHub: **Settings → Pages → Build and deployment → Source:
     GitHub Actions** (configuração única, não precisa repetir)
   - O workflow roda todo dia às 08:00 (horário de Brasília) e publica
     em `https://<seu-usuário>.github.io/<repositório>/` — salve esse
     link na tela inicial do celular
   - **Atenção:** essa página fica pública na internet (sem senha, mas
     ninguém acha sem o link exato); é assim que o GitHub Pages
     funciona no plano gratuito
   - Pra gerar localmente sem esperar o agendamento:
     ```bash
     python atualizar_planilha_financas.py --gerar-painel
     ```
     (gera `painel/index.html`; abra esse arquivo no navegador pra
     conferir antes de publicar)

## Aviso

A extração de PDF e as regras de categorização em `REGRAS` são um ponto de
partida baseado nos extratos do Ourocard Platinum Estilo (BB). Ajuste as
palavras-chave conforme novos comerciantes forem aparecendo nas faturas.
