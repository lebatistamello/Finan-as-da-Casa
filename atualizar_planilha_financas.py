"""
Automação: Fatura do cartão (PDF, salva numa pasta do Google Drive) ->
categorização -> escrita na planilha "Finanças da Família 2026" (Google Sheets).

COMO USAR (via Claude Code):
1. Peça ao Claude Code para revisar este script e instalar as dependências:
   pip install pdfplumber gspread google-auth google-auth-oauthlib \
               google-api-python-client

2. Configure uma Service Account no Google Cloud Console:
   - Crie um projeto (ou use um existente) em console.cloud.google.com
   - Ative a "Google Sheets API" e a "Google Drive API"
   - Crie uma Service Account, gere uma chave JSON, salve como credentials.json
     (NUNCA suba esse arquivo pro GitHub em texto puro — use GitHub Secrets)
   - Compartilhe a planilha "Finanças da Família 2026" COM a pasta de
     faturas no Drive com o e-mail da service account
     (algo como xxxx@yyyy.iam.gserviceaccount.com), dando permissão de Editor
     nos dois

3. Ajuste as constantes no topo do script (SPREADSHEET_ID, DRIVE_FOLDER_ID,
   SHEET_NAME, etc.)

4. Teste manualmente com um PDF local antes de automatizar:
   python atualizar_planilha_financas.py --pdf "fatura_teste.pdf" --mes agosto

5. Teste o modo automático (lê a pasta do Drive) em modo simulação:
   python atualizar_planilha_financas.py --mes agosto

6. Quando validado, use --escrever para gravar de verdade e marcar os
   PDFs como processados:
   python atualizar_planilha_financas.py --mes agosto --escrever

7. Peça ao Claude Code para criar uma Routine (claude.ai/code/routines ou
   /schedule no CLI) apontando pra esse repositório, rodando esse comando
   na frequência que vocês quiserem (semanal / a cada 10 dias) — isso roda
   na nuvem da Anthropic, sem depender de nenhum computador ligado.
   Repare que a Routine vai precisar saber QUAL mês gravar automaticamente
   (hoje o --mes é manual) — vale pedir ao Claude Code pra trocar isso por
   "mês atual" calculado automaticamente pela data de execução.

IMPORTANTE: a extração de PDF e as regras de categorização abaixo são um
PONTO DE PARTIDA baseado nos extratos do Ourocard Platinum Estilo (BB)
analisados nesta conversa. Provavelmente vai precisar de ajustes finos
lançamento a lançamento -- é normal, e o Claude Code pode iterar isso
junto com você olhando saídas reais.
"""

import argparse
import io
import re
import sys
from datetime import date, datetime
from pathlib import Path

import pdfplumber
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ============================================================
# CONFIGURAÇÃO — ajuste antes de rodar
# ============================================================

# ID da planilha "Finanças da Família 2026" (versão nativa Google Sheets,
# convertida a partir do .xlsx original em 15/08/2026 — o Sheets API não
# funciona sobre arquivos .xlsx "modo compatibilidade").
# É o trecho entre /d/ e /edit na URL do Google Sheets.
SPREADSHEET_ID = "1y-HIz6irM70XFt9cJVMK93_flaC_mcgl-0sHNeivObA"

# ID da subpasta do Google Drive onde as faturas em PDF são salvas
# (pasta "Faturas do Cartão da Casa", dentro de "Finanças da Casa",
# no Drive da conta mellmucomunicacao@gmail.com)
DRIVE_FOLDER_ID = "1wnZ0sS_UHeuutQh3wVfDKFc81K9vubAP"

# Nome da aba, confirmado em 15/08/2026 na planilha nativa
SHEET_NAME = "2026"

# Arquivo de credenciais da service account (Sheets API + Drive API)
CREDENTIALS_FILE = "credentials.json"

# Mapeamento mês -> letra da coluna na planilha. Corrigido em 16/08/2026:
# a versão anterior (C..N) estava uma coluna adiantada — confirmado
# contra o texto real do cabeçalho da linha "RECEBIMENTOS" (célula por
# célula, via exportação CSV) e contra a linha "Depósito Lisandra"
# (nunca escrita pelo script, preserva os valores originais do modelo).
# Coluna N é o "budget" (orçamento previsto) de cada categoria.
COLUNA_DO_MES = {
    "janeiro": "B", "fevereiro": "C", "marco": "D", "abril": "E",
    "maio": "F", "junho": "G", "julho": "H", "agosto": "I",
    "setembro": "J", "outubro": "K", "novembro": "L", "dezembro": "M",
}
COLUNA_BUDGET = "N"

# Linha de cada item na planilha. Contado a partir do conteúdo real lido em
# 15/08/2026 (a aba tem, nessa ordem: cabeçalho/objetivos linhas 1-5,
# "O que tenho" 6-13, "O que devo" 14-17, "RECEBIMENTOS" 18-23,
# "DESPESAS FIXAS" a partir da 24). Ainda assim, CONFIRA VISUALMENTE na
# planilha antes de rodar com --escrever. Reconstruído em 16/08/2026 a
# partir de uma exportação CSV bruta da planilha nativa (posição exata de
# cada linha/coluna preservada) — NÃO confiar em leituras de "texto
# natural" do Drive pra isso, elas comprimem/pulam linhas em branco e
# desalinham a contagem (foi exatamente o que causou o erro anterior).
LINHA_DO_ITEM = {
    "Energia elétrica": 33,
    "Água": 34,
    "Gás e Lenha": 35,
    "Internet": 36,
    "Supermercado/feira": 37,
    "Restaurantes/Deliverys": 38,
    "Investimento/Manutenção Casa": 39,
    "Limpeza (Casa e Pátio)": 40,
    "IPTU parcelado 10x": 41,
    "Plano de saúde": 44,
    "Academia/Clube": 45,
    "Farmácia/remédios": 46,
    "Salão de beleza": 47,
    "Atividades Luise": 51,
    "Atividades Maitê": 52,
    "Escola Maitê Marista": 53,
    "Escola Luise Marista": 54,
    "Gasolina CRV": 57,
    "Estacionamento": 58,
    "IPVA Cielo": 59,
    "Seguro CRV": 60,
    "Aplicativos/táxi": 61,
    "Taxas": 63,
    "Assinaturas": 64,
    "PET": 65,
    "Investimentos": 66,
    "Marketplaces": 69,
    "Farmácia (dívida)": 70,
    "Dafitti": 71,
    "Adidas": 72,
    "Outros": 73,
    "Manutenções CRV": 74,
    "Manutenção Cielo": 75,
    "Multas de Trânsito": 78,
    "Compras eventuais": 79,
    "Férias/Viagens": 80,
}

# ============================================================
# REGRAS DE CATEGORIZAÇÃO
# Cada linha de item recebe uma lista de palavras-chave (case-insensitive,
# substring match) que identificam lançamentos daquele tipo no extrato.
# Ajuste/expanda conforme forem aparecendo comerciantes novos.
# ============================================================

REGRAS = {
    "Supermercado/feira": ["SUPER TCHE", "ZAFFARI", "BISTEK", "SAMS CLUB",
                            "MERCADINHO", "BANCA 43", "HORTIFRUTI", "FRUTEIRA"],
    "Restaurantes/Deliverys": ["IFOOD", "RESTAURANT", "PIZZ", "LANCHONETE",
                                "BURGER", "BISTRO", "CAMARADA", "CUNHA E NOSCHANG"],
    "Farmácia/remédios": ["PANVEL", "DROGARIA", "FARMAC"],
    "Academia/Clube": ["ACADEMIA", "AABB"],
    "Salão de beleza": ["ESMALTERIA", "ESTETICA", "SALAO"],
    "Gasolina CRV": ["COMBUSTIVE", "POSTO ", "AUTO POSTO", "ABASTECEDORA",
                      "GAS ZONA SUL", "GASZONASUL"],
    "Seguro CRV": ["VINICIUSGAHBRIEL"],  # corretor do seguro (era da Duster, trocada pela CRV)
    "Estacionamento": ["ESTACIONAMENTO", "HORA PARK", "ALLPARK"],
    "Aplicativos/táxi": ["UBER", "99*", "99 "],
    "Assinaturas": ["SPOTIFY", "NETFLIX", "AMAZON PRIME", "GLOBO PREMIER",
                     "ICLOUD", "YOUTUBE"],
    "PET": ["PET ", "PETSHOP", "VETERINAR"],
    "Escola Maitê Marista": ["ESCOLA MAITE", "MARISTA MAITE"],
    "Escola Luise Marista": ["ESCOLA LUISE", "MARISTA LUISE"],
    "Atividades Maitê": ["IMPULSE"],
    "Taxas": ["TAXA", "IOF", "ANUIDADE"],
    # Regras adicionais vão aparecendo conforme mais faturas forem processadas —
    # peça ao Claude Code pra te ajudar a ir expandindo isso.
}

# Regras que só valem para compras PARCELADAS (descrição com "PARC NN/NN"),
# usadas para alimentar as linhas de "Dívidas/Parcelamentos" da planilha.
# Uma compra à vista nesses mesmos lugares (ex: Mercado Livre à vista) não
# cai aqui — segue as REGRAS normais acima (ou o catch-all de eventuais).
REGRAS_PARCELAMENTO = {
    "Marketplaces": ["MERCADOLIVRE", "MERCADO LIVRE", "MP*MELIMAIS", "MELIMAIS",
                      "AMAZON", "SHOPEE", "ALIEXPRESS", "SHEIN"],
}

# Quando a compra é parcelada, alguns itens das REGRAS normais são
# redirecionados para a linha de dívida equivalente.
REDIRECIONA_SE_PARCELADO = {
    "Farmácia/remédios": "Farmácia (dívida)",
}

PARCELA_REGEX = re.compile(r"PARC\s*\d{1,2}\s*/\s*\d{1,2}", re.IGNORECASE)

DEFAULT_ITEM = "Compras eventuais"  # cai aqui se nada bater

MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "marco", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}

# Abreviações de 3 letras usadas nos nomes de arquivo do app/site do BB
# (ex.: "OUROCARD_PLATINUM_ESTILO_VISA-Abr_26.pdf").
MES_ABREV = {
    "jan": "janeiro", "fev": "fevereiro", "mar": "marco", "abr": "abril",
    "mai": "maio", "jun": "junho", "jul": "julho", "ago": "agosto",
    "set": "setembro", "out": "outubro", "nov": "novembro", "dez": "dezembro",
}


def mes_atual() -> str:
    """Retorna o mês corrente em português, no formato usado por COLUNA_DO_MES.
    Usado pela Routine, que roda sozinha sem ninguém passando --mes na mão."""
    return MESES_PT[date.today().month]


def mes_do_arquivo(nome_arquivo: str):
    """Tenta identificar o mês de referência pelo nome do arquivo (ex.:
    '...-Abr_26.pdf' -> 'abril'). Retorna None se não conseguir identificar,
    para o chamador decidir o que fazer (cair no mês atual, avisar etc.)."""
    nome = nome_arquivo.lower()
    for mes in COLUNA_DO_MES:
        if mes in nome:
            return mes
    for abrev, mes in MES_ABREV.items():
        if re.search(rf"(?<![a-z]){abrev}(?![a-z])", nome):
            return mes
    return None

# ============================================================
# ACESSO AO GOOGLE DRIVE (pasta de faturas)
# ============================================================

def conectar_drive():
    scopes = ["https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    return build("drive", "v3", credentials=creds)


def listar_pdfs_novos(drive_service, folder_id: str):
    """Lista PDFs na pasta que ainda não têm a propriedade 'processado=true',
    do mais antigo pro mais novo. A ordem importa: se duas faturas da mesma
    fatura em aberto (ex: dois downloads semanais de agosto) ficarem
    pendentes ao mesmo tempo, a mais nova precisa ser processada por último
    pra "vencer" — como cada categoria é sobrescrita (não somada), quem
    processa por último decide o valor final da coluna."""
    query = (
        f"'{folder_id}' in parents and mimeType='application/pdf' "
        f"and trashed=false and not properties has {{key='processado' and value='true'}}"
    )
    resultado = drive_service.files().list(
        q=query, fields="files(id, name, createdTime)", pageSize=100,
        orderBy="createdTime"
    ).execute()
    return resultado.get("files", [])


def baixar_pdf_drive(drive_service, file_id: str) -> bytes:
    request = drive_service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buffer.seek(0)
    return buffer.read()


def marcar_como_processado(drive_service, file_id: str):
    drive_service.files().update(
        fileId=file_id,
        body={"properties": {"processado": "true"}},
    ).execute()


def listar_todos_pdfs(drive_service, folder_id: str):
    """Lista TODOS os PDFs da pasta, processados ou não. Usado só por
    --resetar-processados (a interface do Drive não permite editar essa
    propriedade customizada na mão, então esse é o único jeito de desfazer
    uma marcação errada)."""
    query = f"'{folder_id}' in parents and mimeType='application/pdf' and trashed=false"
    resultado = drive_service.files().list(
        q=query, fields="files(id, name)", pageSize=100
    ).execute()
    return resultado.get("files", [])


def desmarcar_processado(drive_service, file_id: str):
    drive_service.files().update(
        fileId=file_id,
        body={"properties": {"processado": None}},
    ).execute()


# ============================================================
# EXTRAÇÃO DO PDF
# ============================================================

def extrair_lancamentos_de_bytes(pdf_bytes: bytes):
    """Mesma extração de extrair_lancamentos, mas a partir de bytes em memória
    (útil quando o PDF vem direto do Drive, sem salvar em disco)."""
    lancamentos = []
    linha_regex = re.compile(
        r"^\s*\d{2}/\d{2}\s+(.+?)\s+(?:BR|[A-Z]{2})?\s*R?\$?\s*([\d.,]+)\s*$"
    )
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                m = linha_regex.match(line)
                if not m:
                    continue
                desc = m.group(1).strip()
                valor_str = m.group(2).replace(".", "").replace(",", ".")
                try:
                    valor = float(valor_str)
                except ValueError:
                    continue
                lancamentos.append((desc, valor))
    return lancamentos


def extrair_lancamentos(pdf_path: str):
    """Extrai (descricao, valor) de todas as linhas de lançamento de um PDF
    local. Mantido para testes manuais com --pdf; o fluxo automático usa
    extrair_lancamentos_de_bytes() a partir do Drive."""
    with open(pdf_path, "rb") as f:
        return extrair_lancamentos_de_bytes(f.read())


def categorizar(descricao: str) -> str:
    desc_upper = descricao.upper()
    parcelado = bool(PARCELA_REGEX.search(descricao))

    if parcelado:
        for item, palavras in REGRAS_PARCELAMENTO.items():
            if any(p in desc_upper for p in palavras):
                return item

    for item, palavras in REGRAS.items():
        if any(p in desc_upper for p in palavras):
            if parcelado and item in REDIRECIONA_SE_PARCELADO:
                return REDIRECIONA_SE_PARCELADO[item]
            return item

    return DEFAULT_ITEM


def somar_por_item(lancamentos):
    totais = {}
    nao_categorizados = []
    for desc, valor in lancamentos:
        item = categorizar(desc)
        totais[item] = totais.get(item, 0.0) + valor
        if item == DEFAULT_ITEM:
            nao_categorizados.append((desc, valor))
    return totais, nao_categorizados


# ============================================================
# ESCRITA NA PLANILHA
# ============================================================

def conectar_planilha():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

def escrever_totais(ws, mes: str, totais: dict, dry_run: bool = True):
    coluna = COLUNA_DO_MES[mes]
    updates = []
    for item, valor in totais.items():
        linha = LINHA_DO_ITEM.get(item)
        if linha is None:
            print(f"[aviso] item '{item}' sem linha mapeada — pulei")
            continue
        cell = f"{coluna}{linha}"
        updates.append((cell, round(valor, 2)))

    print(f"\n{'[SIMULAÇÃO] ' if dry_run else ''}Valores que {'seriam' if dry_run else 'foram'} escritos (coluna {coluna}):")
    for cell, valor in updates:
        print(f"  {cell} = R$ {valor:,.2f}")

    if not dry_run:
        for cell, valor in updates:
            ws.update_acell(cell, valor)
        print("\nPlanilha atualizada.")
    else:
        print("\n(Rodando em modo simulação — use --escrever para gravar de verdade)")


# ============================================================
# PAINEL (mês atual x orçamento) — página HTML pra acompanhar pelo celular
# ============================================================

def parse_valor_br(texto) -> float:
    """Converte 'R$ 3.000,00' / '396,90' / '' / None em float. Célula vazia
    ou não numérica vira 0.0 (linha ainda sem orçamento/gasto definido)."""
    texto = (texto or "").strip()
    if not texto:
        return 0.0
    texto = texto.replace("R$", "").strip()
    texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def fmt_brl(valor: float) -> str:
    s = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def _col_para_indice(letra: str) -> int:
    """Converte letra de coluna (ex: 'E', 'N') pra índice 0-based."""
    n = 0
    for ch in letra.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def gerar_painel_html(ws, mes: str) -> str:
    """Lê o orçamento (coluna budget) e o gasto do mês atual (coluna do mês)
    de cada categoria em LINHA_DO_ITEM e monta uma página HTML simples,
    lado a lado, com a diferença (orçamento - gasto).

    Busca a planilha inteira numa única chamada (get_all_values) em vez de
    uma chamada por célula — com ~35 categorias, ler célula a célula bate
    fácil na cota de "leituras por minuto" da API do Sheets."""
    coluna_atual = COLUNA_DO_MES[mes]
    idx_budget = _col_para_indice(COLUNA_BUDGET)
    idx_atual = _col_para_indice(coluna_atual)
    valores = ws.get_all_values()

    def valor_em(linha: int, idx_col: int) -> str:
        linha_dados = valores[linha - 1] if linha - 1 < len(valores) else []
        return linha_dados[idx_col] if idx_col < len(linha_dados) else ""

    linhas_html = []
    total_budget = 0.0
    total_atual = 0.0

    for item, linha in LINHA_DO_ITEM.items():
        budget = parse_valor_br(valor_em(linha, idx_budget))
        atual = parse_valor_br(valor_em(linha, idx_atual))
        if budget == 0 and atual == 0:
            continue  # categoria sem orçamento e sem gasto neste mês — não polui o painel
        diff = budget - atual
        total_budget += budget
        total_atual += atual
        cor = "#1a7f37" if diff >= 0 else "#cf222e"
        linhas_html.append(
            f"<tr><td>{item}</td>"
            f"<td class='num'>{fmt_brl(budget)}</td>"
            f"<td class='num'>{fmt_brl(atual)}</td>"
            f"<td class='num' style='color:{cor}'>{fmt_brl(diff)}</td></tr>"
        )

    diff_total = total_budget - total_atual
    cor_total = "#1a7f37" if diff_total >= 0 else "#cf222e"
    atualizado_em = datetime.now().strftime("%d/%m/%Y %H:%M")

    return f"""<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Orçamento — {mes.capitalize()}</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><text y=%22.9em%22 font-size=%2290%22>%F0%9F%8F%A0</text></svg>">
<link rel="apple-touch-icon" href="data:image/svg+xml,<svg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 100 100%22><rect width=%22100%22 height=%22100%22 rx=%2220%22 fill=%22%23ffffff%22/><text x=%2250%22 y=%2270%22 font-size=%2260%22 text-anchor=%22middle%22>%F0%9F%8F%A0</text></svg>">
<style>
  :root {{ color-scheme: light dark; }}
  body {{ font-family: -apple-system, system-ui, sans-serif; margin: 0;
          padding: 16px; background: #ffffff; color: #1f2328; }}
  @media (prefers-color-scheme: dark) {{
    body {{ background: #0d1117; color: #e6edf3; }}
    th {{ background: #161b22 !important; }}
    tr:nth-child(even) {{ background: #161b22; }}
    tfoot td {{ border-top-color: #30363d !important; }}
  }}
  h1 {{ font-size: 1.3rem; margin: 0 0 2px; }}
  .atualizado {{ font-size: 0.8rem; opacity: 0.65; margin: 0 0 16px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.92rem; }}
  th, td {{ padding: 8px 6px; text-align: left; }}
  th {{ background: #f0f2f5; font-size: 0.78rem; text-transform: uppercase;
        letter-spacing: 0.02em; }}
  td.num, th.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
  tr:nth-child(even) {{ background: #f6f8fa; }}
  tfoot td {{ font-weight: 700; border-top: 2px solid #d0d7de; padding-top: 10px; }}
</style>
</head>
<body>
  <h1>Orçamento de {mes.capitalize()}</h1>
  <p class="atualizado">Atualizado em {atualizado_em}</p>
  <table>
    <thead>
      <tr><th>Categoria</th><th class="num">Orçamento</th><th class="num">Gasto</th><th class="num">Diferença</th></tr>
    </thead>
    <tbody>
      {''.join(linhas_html)}
    </tbody>
    <tfoot>
      <tr><td>Total</td>
        <td class="num">{fmt_brl(total_budget)}</td>
        <td class="num">{fmt_brl(total_atual)}</td>
        <td class="num" style="color:{cor_total}">{fmt_brl(diff_total)}</td>
      </tr>
    </tfoot>
  </table>
</body>
</html>
"""


# ============================================================
# MAIN
# ============================================================

def processar_e_escrever(ws, mes: str, lancamentos, escrever: bool):
    """Categoriza os lançamentos de UMA fatura e grava (ou simula) os totais
    na coluna do mês correspondente."""
    print(f"{len(lancamentos)} lançamentos encontrados.")

    totais, nao_categorizados = somar_por_item(lancamentos)

    print("\nTotais por item:")
    for item, valor in sorted(totais.items(), key=lambda x: -x[1]):
        print(f"  {item}: R$ {valor:,.2f}")

    if nao_categorizados:
        print(f"\n[atenção] {len(nao_categorizados)} lançamentos caíram em "
              f"'{DEFAULT_ITEM}' por falta de regra — confira se fazem sentido:")
        for desc, valor in nao_categorizados[:20]:
            print(f"  - {desc}: R$ {valor:,.2f}")

    escrever_totais(ws, mes, totais, dry_run=not escrever)


def main():
    parser = argparse.ArgumentParser(description="Processa fatura(s) do cartão e atualiza a planilha oficial.")
    parser.add_argument("--pdf", help="Caminho de um PDF local (modo manual/teste)")
    parser.add_argument("--mes", choices=list(COLUNA_DO_MES.keys()), default=None,
                         help="Mês de referência (ex: agosto). No modo --pdf, se omitido usa o mês "
                              "atual. No modo automático (pasta do Drive), se omitido cada fatura "
                              "tem o mês identificado pelo próprio nome do arquivo (ex: "
                              "'...-Abr_26.pdf' -> abril); passar --mes força esse mês pra TODAS "
                              "as faturas encontradas na pasta.")
    parser.add_argument("--escrever", action="store_true",
                         help="Grava de verdade na planilha (padrão: só simula)")
    parser.add_argument("--resetar-processados", action="store_true",
                         help="Desmarca TODOS os PDFs da pasta como não processados (usado pra "
                              "corrigir uma gravação errada e permitir reprocessar as mesmas "
                              "faturas). Não mexe na planilha, só no Drive.")
    parser.add_argument("--gerar-painel", action="store_true",
                         help="Gera o painel HTML (mês atual x orçamento) em --painel-saida e "
                              "sai, sem processar faturas. Só lê a planilha.")
    parser.add_argument("--painel-saida", default="painel/index.html",
                         help="Caminho do arquivo HTML gerado por --gerar-painel "
                              "(padrão: painel/index.html)")
    args = parser.parse_args()

    if args.resetar_processados:
        drive = conectar_drive()
        pdfs = listar_todos_pdfs(drive, DRIVE_FOLDER_ID)
        if not pdfs:
            print("Nenhum PDF encontrado na pasta do Drive.")
            return
        for f in pdfs:
            desmarcar_processado(drive, f["id"])
            print(f"'{f['name']}' desmarcado como processado.")
        return

    ws = conectar_planilha()

    if args.gerar_painel:
        mes = args.mes or mes_atual()
        html = gerar_painel_html(ws, mes)
        saida = Path(args.painel_saida)
        saida.parent.mkdir(parents=True, exist_ok=True)
        saida.write_text(html, encoding="utf-8")
        print(f"Painel gerado em {saida} (mês: {mes})")
        return

    if args.pdf:
        # Modo manual: um PDF local, pra teste
        mes = args.mes or mes_atual()
        print(f"Mês de referência: {mes}" + (" (detectado automaticamente)" if not args.mes else ""))
        pdf_path = Path(args.pdf)
        if not pdf_path.exists():
            sys.exit(f"Arquivo não encontrado: {pdf_path}")
        print(f"Lendo {pdf_path.name} (local)...")
        lancamentos = extrair_lancamentos(str(pdf_path))
        processar_e_escrever(ws, mes, lancamentos, args.escrever)
        return

    # Modo automático: varre a pasta do Drive por PDFs ainda não processados,
    # cada um gravado no mês que lhe corresponde.
    drive = conectar_drive()
    pdfs_novos = listar_pdfs_novos(drive, DRIVE_FOLDER_ID)
    if not pdfs_novos:
        print("Nenhuma fatura nova encontrada na pasta do Drive.")
        return
    print(f"{len(pdfs_novos)} fatura(s) nova(s) encontrada(s) na pasta do Drive:")
    for f in pdfs_novos:
        print(f"  - {f['name']}")

    for f in pdfs_novos:
        if args.mes:
            mes = args.mes
        else:
            mes = mes_do_arquivo(f["name"])
            if mes is None:
                mes = mes_atual()
                print(f"\n[aviso] não identifiquei o mês pelo nome de '{f['name']}' "
                      f"— usando o mês atual ({mes})")

        print(f"\n--- {f['name']} -> mês: {mes} ---")
        pdf_bytes = baixar_pdf_drive(drive, f["id"])
        lancamentos = extrair_lancamentos_de_bytes(pdf_bytes)
        processar_e_escrever(ws, mes, lancamentos, args.escrever)

        if args.escrever:
            marcar_como_processado(drive, f["id"])
            print(f"'{f['name']}' marcado como processado no Drive.")


if __name__ == "__main__":
    main()
