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

# Linha de cada item na planilha. A família removeu o bloco "O que tenho" /
# "O que devo" (17 linhas) que existia acima de RECEBIMENTOS quando
# reestruturou a aba em 27/08/2026 — isso empurrou TODAS as linhas abaixo 24
# posições pra cima, e como os números aqui não foram reajustados, toda
# gravação de --escrever entre 27/08 e 30/08 escreveu 24 linhas abaixo de
# onde deveria (em linhas em branco, ou pior, por cima de outra categoria/de
# fórmulas como SALDO DO MÊS). Renumerado em 30/08/2026 direto de uma
# exportação CSV bruta da planilha nativa (posição exata de cada linha/coluna
# preservada) — NÃO confiar em leituras de "texto natural" do Drive pra isso,
# elas comprimem/pulam linhas em branco e desalinham a contagem. Ainda assim,
# CONFIRA VISUALMENTE na planilha antes de rodar com --escrever.
LINHA_DO_ITEM = {
    "Energia elétrica": 9,
    "Água": 10,
    "Gás e Lenha": 11,
    "Internet": 12,
    "Supermercado/feira": 13,
    "Restaurantes/Deliverys": 14,
    "Investimento/Manutenção Casa": 15,
    "Limpeza (Casa e Pátio)": 16,
    "IPTU parcelado 10x": 17,
    "Plano de saúde": 20,
    "Academia/Clube": 21,
    "Farmácia/remédios": 22,
    "Salão de beleza": 23,
    "Atividades Luise": 27,
    "Atividades Maitê": 28,
    "Escola Maitê Marista": 29,
    "Escola Luise Marista": 30,
    "Gasolina CRV": 33,
    "Pedágio/Estacionamento": 34,
    "IPVA Cielo": 35,
    "Seguro CRV": 36,
    "Aplicativos/táxi": 37,
    "Taxas": 39,
    "Assinaturas": 40,
    "PET": 41,
    "Investimentos": 42,
    "Mercado Livre": 45,
    "Farmácia (dívida)": 46,
    "Dafiti": 47,
    "Adidas": 48,
    "Outros parcelamentos": 49,
    "Manutenções CRV": 50,
    "Manutenção Cielo": 51,
    "Multas de Trânsito": 54,
    "Compras eventuais à vista": 55,
    "Férias/Viagens": 56,
}

# ============================================================
# REGRAS DE CATEGORIZAÇÃO
# Cada linha de item recebe uma lista de palavras-chave (case-insensitive,
# substring match) que identificam lançamentos daquele tipo no extrato.
# Ajuste/expanda conforme forem aparecendo comerciantes novos.
# ============================================================

REGRAS = {
    "Supermercado/feira": ["SUPER TCHE", "ZAFFARI", "BISTEK", "SAMS CLUB",
                            "MERCADINHO", "BANCA 43", "HORTIFRUTI", "FRUTEIRA",
                            "SHOPPING DE CARNES"],
    "Restaurantes/Deliverys": ["IFOOD", "RESTAURANT", "PIZZ", "LANCHONETE",
                                "BURGER", "BISTRO", "CAMARADA", "CUNHA E NOSCHANG"],
    "Farmácia/remédios": ["PANVEL", "DROGARIA", "FARMAC", "RAIA", "DROGA RAIA",
                           "FARMACIAS SAO JOAO", "FARMACIA SAO JOAO", "SAO JOAO FARM",
                           "PAGUE MENOS", "DROGASIL"],
    "Academia/Clube": ["ACADEMIA", "AABB"],
    "Salão de beleza": ["ESMALTERIA", "ESTETICA", "SALAO"],
    "Investimento/Manutenção Casa": ["ROBERTA BALESTRIN", "CASSOL", "FERRAGEM",
                                      "MATERIAL DE CONSTRUCAO", "TINTAS", "LEROY MERLIN",
                                      "TELHANORTE", "C&C CASA", "MARCENARIA", "SERRALHERIA"],
    "Gasolina CRV": ["COMBUSTIVE", "POSTO ", "AUTO POSTO", "ABASTECEDORA",
                      "GAS ZONA SUL", "GASZONASUL"],
    "Seguro CRV": ["VINICIUSGAHBRIEL"],  # corretor do seguro (era da Duster, trocada pela CRV)
    "Pedágio/Estacionamento": ["ESTACIONAMENTO", "HORA PARK", "ALLPARK", "ESTAPAR", "ZUL PARK",
                                "SEM PARAR", "CONECTCAR", "VELOE", "MOVE MAIS", "TAGGY", "PEDAGIO"],
    "Aplicativos/táxi": ["UBER", "99*", "99 "],
    "Assinaturas": ["SPOTIFY", "NETFLIX", "AMAZON PRIME", "AMAZONPRIME", "GLOBO PREMIER",
                     "ICLOUD", "YOUTUBE"],
    "PET": ["PET ", "PETSHOP", "VETERINAR", "COBASI"],
    "Escola Maitê Marista": ["ESCOLA MAITE", "MARISTA MAITE"],
    "Escola Luise Marista": ["ESCOLA LUISE", "MARISTA LUISE"],
    "Atividades Maitê": ["IMPULSE"],
    "Taxas": ["TAXA", "IOF", "ANUIDADE"],
    "Dafiti": ["DAFITI"],
    "Adidas": ["ADIDAS"],
    "Mercado Livre": ["MERCADOLIVRE", "MERCADO LIVRE", "MP*MELIMAIS", "MELIMAIS",
                       "AMAZON", "SHOPEE", "ALIEXPRESS", "SHEIN"],
    # Regras adicionais vão aparecendo conforme mais faturas forem processadas —
    # peça ao Claude Code pra te ajudar a ir expandindo isso.
}
# Todo comerciante mapeado acima cai sempre na mesma linha, à vista ou
# parcelado (confirmado com a família em 27/08/2026) — só o que NÃO bate
# com nenhuma regra é que se divide entre "à vista" e "parcelado" (ver
# categorizar() e ITEM_PARCELAMENTO_GENERICO/DEFAULT_ITEM logo abaixo).

PARCELA_REGEX = re.compile(r"PARC\s*\d{1,2}\s*/\s*\d{1,2}", re.IGNORECASE)

DEFAULT_ITEM = "Compras eventuais à vista"  # cai aqui se nada bater e não for parcelado
ITEM_PARCELAMENTO_GENERICO = "Outros parcelamentos"  # cai aqui se for parcelado (PARC NN/NN)
                                                       # mas não bateu com nenhum comerciante
                                                       # conhecido — dentro de Dívidas/Parcelamentos,
                                                       # não misturado com compras avulsas do dia a dia

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
    # Sem âncora de fim de linha (nem $): pega o PRIMEIRO valor em R$ depois da
    # descrição, não o último. O layout dessa fatura às vezes emenda duas
    # transações na mesma linha de texto (quando caem na mesma altura/y da
    # página) ou emenda um "Total" de seção logo depois do último lançamento —
    # com âncora de fim de linha, a regra antiga pegava esse valor errado (o
    # da direita) e engolia a transação de verdade inteira dentro da descrição.
    # finditer (em vez de match) faz o mesmo scan pegar as DUAS transações
    # quando elas vêm emendadas, ao invés de só reconhecer a primeira e
    # descartar a segunda.
    linha_regex = re.compile(
        r"(?:^|\s)\d{2}/\d{2}\s+(.+?)\s+(?:BR|[A-Z]{2})?\s*R?\$?\s*(\d{1,3}(?:\.\d{3})*,\d{2})\b"
    )
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                for m in linha_regex.finditer(line):
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

    for item, palavras in REGRAS.items():
        if any(p in desc_upper for p in palavras):
            return item

    # não bateu com nenhum comerciante conhecido — mas se o texto tem "PARC 01/03"
    # (padrão de parcelamento do Ourocard), é uma dívida/parcelamento não mapeado
    # ainda, não uma compra eventual comum
    if PARCELA_REGEX.search(descricao):
        return ITEM_PARCELAMENTO_GENERICO
    return DEFAULT_ITEM


def somar_por_item(lancamentos):
    totais = {}
    nao_categorizados = []
    for desc, valor in lancamentos:
        item = categorizar(desc)
        totais[item] = totais.get(item, 0.0) + valor
        if item in (DEFAULT_ITEM, ITEM_PARCELAMENTO_GENERICO):
            nao_categorizados.append((desc, valor))
    return totais, nao_categorizados


# ============================================================
# ESCRITA NA PLANILHA
# ============================================================

def conectar_planilha():
    """Retorna (spreadsheet, worksheet_principal). O spreadsheet é necessário
    à parte pra poder criar/acessar as abas de detalhamento (ver
    escrever_detalhamento)."""
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    return spreadsheet, spreadsheet.worksheet(SHEET_NAME)

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
# ABAS DE DETALHAMENTO — "Compras eventuais à vista" e "Outros parcelamentos"
#
# Toda vez que uma fatura é processada, a aba com o nome exato da categoria
# é SUBSTITUÍDA por inteiro com a lista de lançamentos que caíram nela NAQUELA
# fatura (não é um acumulado histórico). Como o BB atualiza a mesma fatura em
# aberto conforme o mês avança, reprocessar semanalmente já mantém a aba
# refletindo o mês inteiro até a data do processamento — mesma lógica de
# "sobrescreve, não soma" usada pros totais na aba principal.
# ============================================================

ITENS_DETALHADOS = [DEFAULT_ITEM, ITEM_PARCELAMENTO_GENERICO]


def obter_ou_criar_aba(spreadsheet, nome: str):
    try:
        return spreadsheet.worksheet(nome)
    except gspread.exceptions.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=nome, rows=200, cols=4)


def escrever_detalhamento(spreadsheet, mes: str, lancamentos: list, dry_run: bool = True):
    for item in ITENS_DETALHADOS:
        detalhes = [(desc, valor) for desc, valor in lancamentos if categorizar(desc) == item]
        detalhes.sort(key=lambda x: -x[1])  # maiores gastos primeiro

        print(f"\n{'[SIMULAÇÃO] ' if dry_run else ''}Detalhamento de '{item}' ({mes}): "
              f"{len(detalhes)} lançamento(s), total R$ {sum(v for _, v in detalhes):,.2f}")
        for desc, valor in detalhes[:15]:
            print(f"  {desc}: R$ {valor:,.2f}")
        if len(detalhes) > 15:
            print(f"  ... e mais {len(detalhes) - 15} lançamento(s)")

        if dry_run:
            continue

        ws_detalhe = obter_ou_criar_aba(spreadsheet, item)
        ws_detalhe.clear()
        linhas = [[f"{item} — {mes.capitalize()}/2026", "", ""],
                  ["Descrição", "Valor (R$)", ""]]
        for desc, valor in detalhes:
            linhas.append([desc, round(valor, 2), ""])
        linhas.append(["TOTAL", round(sum(v for _, v in detalhes), 2), ""])
        ws_detalhe.update(linhas, "A1")
        print(f"Aba '{item}' substituída com {len(detalhes)} lançamento(s).")


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
        total_budget += budget
        total_atual += atual

        pct = (atual / budget * 100) if budget > 0 else (100.0 if atual > 0 else 0.0)
        pct_barra = min(pct, 100.0)
        diferenca = atual - budget
        if budget <= 0:
            status = "estourado" if atual > 0 else "good"
        elif diferenca > 0.005:
            status = "estourado"  # passou do orçado — roxo
        elif abs(diferenca) <= 0.005:
            status = "critical"  # bateu certinho no orçado — vermelho
        elif pct >= 70:
            status = "warning"
        else:
            status = "good"

        estouro_html = ""
        if atual > budget > 0:
            estouro_html = (
                f'<p class="estouro">⚠ estourou em {fmt_brl(diferenca)}</p>'
            )

        # Quanto ainda falta pra bater no orçado — só faz sentido mostrar em
        # good/warning (em "critical"/"estourado" a sobra é zero ou negativa).
        # Mostra dentro da própria barra (no trilho, depois do preenchimento)
        # quando sobra espaço de sobra pra caber o texto sem cortar; a
        # estimativa é grosseira (não temos as larguras reais renderizadas
        # aqui no servidor), então em barras muito cheias joga a mesma
        # informação pra uma linha abaixo em vez de arriscar cortar o texto
        # dentro do trilho.
        disponivel_dentro = ""
        disponivel_fora = ""
        if status in ("good", "warning") and budget > 0:
            disponivel = budget - atual
            valor_disponivel = fmt_brl(disponivel)
            trilho_px_estimado = 190
            espaco_livre_px = trilho_px_estimado * (100 - pct_barra) / 100
            # Versão de dentro do trilho fica só com o valor (sem a palavra
            # "disponível") pra caber em mais barras — a posição já deixa
            # claro que é a sobra. De fora (fallback), o texto ganha a
            # palavra de volta porque perde esse contexto posicional.
            texto_px_estimado = len(valor_disponivel) * 6.5 + 12
            if espaco_livre_px >= texto_px_estimado:
                disponivel_dentro = f'<span class="disponivel">{valor_disponivel}</span>'
            else:
                disponivel_fora = f'<p class="disponivel-fora">{valor_disponivel} disponível</p>'

        linhas_html.append(f"""
      <div class="item">
        <p class="categoria">{item}</p>
        <div class="medidor-linha">
          <span class="valor valor-gasto">{fmt_brl(atual)}</span>
          <div class="trilho">
            <div class="preenchimento {status}" style="width:{pct_barra:.1f}%"></div>
            {disponivel_dentro}
          </div>
          <span class="valor valor-orcamento">{fmt_brl(budget)}</span>
        </div>
        {estouro_html}
        {disponivel_fora}
      </div>""")

    diff_total = total_budget - total_atual
    pct_total = (total_atual / total_budget * 100) if total_budget > 0 else 0.0
    cor_total = "var(--good)" if diff_total >= 0 else "var(--critical)"
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
  :root {{
    color-scheme: light;
    --surface: #fcfcfb;
    --text-primary: #0b0b0b;
    --text-secondary: #52514e;
    --text-muted: #898781;
    --trilho: #e1e0d9;
    --good: #0ca30c;
    --warning: #fab219;
    --critical: #d03b3b;
    --estourado: #4a3aa7;
    --border: rgba(11,11,11,0.10);
  }}
  @media (prefers-color-scheme: dark) {{
    :root {{
      color-scheme: dark;
      --surface: #1a1a19;
      --text-primary: #ffffff;
      --text-secondary: #c3c2b7;
      --text-muted: #898781;
      --trilho: #2c2c2a;
      --good: #0ca30c;
      --warning: #fab219;
      --critical: #e66767;
      --estourado: #9085e9;
      --border: rgba(255,255,255,0.10);
    }}
  }}
  body {{
    font-family: -apple-system, system-ui, "Segoe UI", sans-serif; margin: 0;
    padding: 16px; background: var(--surface); color: var(--text-primary);
    overflow-x: hidden;
  }}
  .item {{ max-width: 100%; }}
  h1 {{ font-size: 1.3rem; margin: 0 0 2px; }}
  .atualizado {{ font-size: 0.8rem; color: var(--text-muted); margin: 0 0 4px; }}
  .resumo {{ font-size: 0.95rem; color: var(--text-secondary); margin: 0 0 20px; }}
  .resumo strong {{ color: {cor_total}; }}
  .item {{ padding: 10px 0; border-bottom: 1px solid var(--border); }}
  .item:last-child {{ border-bottom: none; }}
  .categoria {{ margin: 0 0 6px; font-size: 0.92rem; }}
  .medidor-linha {{
    display: grid; grid-template-columns: auto minmax(0, 1fr) auto;
    align-items: center; gap: 8px; width: 100%;
  }}
  .valor {{
    font-size: 0.8rem; font-variant-numeric: tabular-nums; color: var(--text-secondary);
    white-space: nowrap;
  }}
  .valor-gasto {{ color: var(--text-primary); font-weight: 600; }}
  .trilho {{
    position: relative; min-width: 0; height: 20px; border-radius: 999px;
    background: var(--trilho); overflow: hidden;
  }}
  .preenchimento {{ height: 100%; border-radius: 999px; }}
  .preenchimento.good {{ background: var(--good); }}
  .preenchimento.warning {{ background: var(--warning); }}
  .preenchimento.critical {{ background: var(--critical); }}
  .preenchimento.estourado {{ background: var(--estourado); }}
  .disponivel {{
    position: absolute; top: 50%; right: 8px; transform: translateY(-50%);
    font-size: 0.66rem; font-variant-numeric: tabular-nums; color: var(--text-secondary);
    white-space: nowrap;
  }}
  .disponivel-fora {{ margin: 4px 0 0; font-size: 0.78rem; color: var(--text-muted); }}
  .estouro {{ margin: 4px 0 0; font-size: 0.78rem; color: var(--estourado); }}
</style>
</head>
<body>
  <h1>Orçamento de {mes.capitalize()}</h1>
  <p class="atualizado">Atualizado em {atualizado_em}</p>
  <p class="resumo">Total: <strong>{fmt_brl(total_atual)}</strong> de {fmt_brl(total_budget)} previstos ({pct_total:.0f}%)</p>
  <div class="lista">
    {''.join(linhas_html)}
  </div>
</body>
</html>
"""


# ============================================================
# MAIN
# ============================================================

def processar_e_escrever(spreadsheet, ws, mes: str, lancamentos, escrever: bool):
    """Categoriza os lançamentos de UMA fatura, grava (ou simula) os totais
    na coluna do mês correspondente, e atualiza as abas de detalhamento das
    categorias "catch-all" (Compras eventuais à vista / Outros parcelamentos)."""
    print(f"{len(lancamentos)} lançamentos encontrados.")

    totais, nao_categorizados = somar_por_item(lancamentos)

    print("\nTotais por item:")
    for item, valor in sorted(totais.items(), key=lambda x: -x[1]):
        print(f"  {item}: R$ {valor:,.2f}")

    if nao_categorizados:
        print(f"\n[atenção] {len(nao_categorizados)} lançamentos caíram em "
              f"'{DEFAULT_ITEM}' ou '{ITEM_PARCELAMENTO_GENERICO}' por falta de "
              f"regra — confira se fazem sentido:")
        for desc, valor in nao_categorizados[:20]:
            print(f"  - {desc}: R$ {valor:,.2f}")

    escrever_totais(ws, mes, totais, dry_run=not escrever)
    escrever_detalhamento(spreadsheet, mes, lancamentos, dry_run=not escrever)


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
    parser.add_argument("--limpar-celulas",
                         help="Manutenção: apaga o conteúdo das células informadas (ex: "
                              "'I57,I58') e sai, sem processar faturas. Usado pra corrigir "
                              "lixo deixado por uma gravação com LINHA_DO_ITEM desalinhado.")
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

    if args.limpar_celulas:
        _, ws = conectar_planilha()
        celulas = [c.strip() for c in args.limpar_celulas.split(",") if c.strip()]
        for cell in celulas:
            ws.update_acell(cell, "")
            print(f"Célula {cell} limpa.")
        return

    spreadsheet, ws = conectar_planilha()

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
        processar_e_escrever(spreadsheet, ws, mes, lancamentos, args.escrever)
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
        processar_e_escrever(spreadsheet, ws, mes, lancamentos, args.escrever)

        if args.escrever:
            marcar_como_processado(drive, f["id"])
            print(f"'{f['name']}' marcado como processado no Drive.")


if __name__ == "__main__":
    main()
