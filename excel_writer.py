import shutil
import openpyxl
from pathlib import Path
from datetime import date
from config import MODELO_FILE, OUTPUT_DIR

MESES_PT = {
    1: "JANEIRO",  2: "FEVEREIRO", 3: "MARÇO",    4: "ABRIL",
    5: "MAIO",     6: "JUNHO",     7: "JULHO",     8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO",  11: "NOVEMBRO", 12: "DEZEMBRO"
}

def extrair_mes_do_periodo(periodo: str) -> str:
    """
    "01/02/2026 A 28/02/2026" → "FEVEREIRO 2026"
    "01/02 A 28/02/2026"      → "FEVEREIRO 2026"
    "01/02 A 28/02"           → "FEVEREIRO"
    """
    try:
        partes = periodo.upper().strip().split("A")
        data_fim = partes[-1].strip()
        segmentos = data_fim.replace("-", "/").split("/")
        if len(segmentos) >= 3:
            mes = int(segmentos[1])
            ano = segmentos[2][:4]
            return f"{MESES_PT.get(mes, '')} {ano}"
        elif len(segmentos) == 2:
            mes = int(segmentos[1])
            return MESES_PT.get(mes, "")
    except Exception:
        pass
    return ""


def _fmt_date(d) -> str:
    if isinstance(d, date):
        return d.strftime("%d/%m/%Y")
    try:
        import pandas as pd
        return pd.to_datetime(d).strftime("%d/%m/%Y")
    except Exception:
        return str(d)


def _w(ws, cell_coord: str, value):
    """
    Escreve valor na célula SEM alterar nenhuma formatação existente.
    Preserva fonte, borda, preenchimento, alinhamento e formato numérico.
    Assume que 'cell_coord' é a célula superior esquerda de qualquer mesclagem relevante.
    """
    ws[cell_coord].value = value


def _wn(ws, row: int, col: int, value):
    """Escreve por índice numérico sem alterar formatação."""
    cell_coord = openpyxl.utils.cell.get_column_letter(col) + str(row)
    _w(ws, cell_coord, value)


def garantir_modelo_no_repo():
    """
    Verifica se o modelo existe. Se não existir, lança erro com
    instrução clara para o usuário colocar o arquivo no repositório.
    """
    if not MODELO_FILE.exists():
        raise FileNotFoundError(
            f"Arquivo modelo não encontrado em: {MODELO_FILE}\n"
            f"Certifique-se de que o arquivo 'Modelo_medio.xlsx' está "
            f"na pasta 'modelo/' do repositório."
        )


def gerar_excel_medicao(dados: dict) -> Path:
    """
    Copia o modelo original e preenche APENAS as células variáveis.
    Todo o resto (formatação, estrutura, textos fixos) é preservado.
    """

    garantir_modelo_no_repo()

    # ── Desempacota dados ─────────────────────────────────────
    contrato          = str(dados["contrato"])
    num_medicao       = int(dados["num_medicao"])
    periodo           = str(dados["periodo"])
    data_apresentacao = dados["data_apresentacao"]
    descricao_servico = str(dados["descricao_servico"])
    valor_mes         = float(dados["valor_mes"])
    quant_mes         = float(dados["quant_mes"])

    empresa           = str(dados["empresa"])
    local             = str(dados["local"])
    modalidade        = str(dados["modalidade"])
    data_base         = dados["data_base"]
    data_termino      = dados["data_termino"]
    valor_original    = float(dados["valor_original"])
    item_num          = dados["item_num"]
    und               = str(dados["und"])
    quant_total       = float(dados["quant_total"])
    preco_unitario    = float(dados["preco_unitario"])
    centro_custo      = str(dados.get("centro_custo", ""))
    conta_contabil    = str(dados.get("conta_contabil", ""))
    item_caixa        = str(dados.get("item_caixa", ""))

    quant_acum_ant    = float(dados["quant_acum_ant"])
    valor_acum_ant    = float(dados["valor_acum_ant"])
    quant_acum_total  = float(dados["quant_acum_total"])
    valor_acum_total  = float(dados["valor_acum_total"])

    # Lista de boletins: [{"label": str, "valor": float}, ...]
    historico         = dados.get("historico", [])

    mes_execucao = extrair_mes_do_periodo(periodo)

    # ── Nome do arquivo de saída ──────────────────────────────
    mes_slug = mes_execucao.replace(" ", "_")
    nome = f"medicao_{num_medicao:02d}_{contrato}_{mes_slug}.xlsx"
    output_path = OUTPUT_DIR / nome

    # Copia o modelo — preserva TODA a formatação original
    shutil.copy(MODELO_FILE, output_path)
    wb = openpyxl.load_workbook(output_path)

    # Nome exato das abas conforme o arquivo modelo
    NOME_PROTOCOLO = "02-26 PROTOCOLO"
    NOME_BOLETIM   = "02-26 BOLETIM"

    # ══════════════════════════════════════════════════════════
    # ABA PROTOCOLO
    # ══════════════════════════════════════════════════════════
    ws = wb[NOME_PROTOCOLO]

    # ── Linha 6: número do boletim ────────────────────────────
    _w(ws, "A6", f"Boletim de Medição n. {num_medicao:02d}")

    # ── Linha 7: descrição do serviço (escolhida na lista) ────
    _w(ws, "A7", descricao_servico)

    # ── Dados fixos do contrato (verde) ───────────────────────
    _w(ws, "D3",  contrato) # Célula superior esquerda da mesclagem D3:F3
    _w(ws, "B8",  empresa)
    _w(ws, "B9",  contrato)
    _w(ws, "H9",  modalidade)
    _w(ws, "B10", local)
    _w(ws, "B11", _fmt_date(data_base))
    _w(ws, "H12", _fmt_date(data_termino))
    _w(ws, "H17", valor_original)
    _w(ws, "H18", valor_original)
    _w(ws, "H23", valor_original)
    _w(ws, "H25", valor_original)

    # Centro de custo / conta contábil / item caixa (linha 63)
    # A63 é a célula superior esquerda da mesclagem A63:B63
    _w(ws, "A63", f"CENTRO DE CUSTO: {centro_custo}")
    # C63 é a célula superior esquerda da mesclagem C63:E63
    _w(ws, "C63", f"CONTA CONTÁBIL: {conta_contabil}")
    # F63 é a célula superior esquerda da mesclagem F63:G63
    _w(ws, "F63", f"ITEM CAIXA: {item_caixa}")

    # ── Dados variáveis da medição (vermelho) ─────────────────
    _w(ws, "B12", periodo)
    _w(ws, "C14", mes_execucao) # C14 é a célula superior esquerda da mesclagem C14:F14
    _w(ws, "C15", _fmt_date(data_apresentacao)) # C15 é a célula superior esquerda da mesclagem C15:F15

    # ── Bloco amarelo: histórico cumulativo ───────────────────
    # Linhas 30 a 44 no modelo (15 boletins)
    # A cada nova medição, limpamos e reescrevemos tudo
    LINHA_HIST_INICIO = 30

    # Apaga bloco inteiro (até 100 linhas para segurança)
    for r in range(LINHA_HIST_INICIO, LINHA_HIST_INICIO + 100):
        ws.cell(row=r, column=4).value = None  # coluna D
        ws.cell(row=r, column=8).value = None  # coluna H

    # Escreve histórico (um boletim por linha)
    for i, h in enumerate(historico):
        row = LINHA_HIST_INICIO + i
        # D30 é a célula superior esquerda da mesclagem D30:G30
        ws.cell(row=row, column=4).value = h["label"]
        ws.cell(row=row, column=8).value = h["valor"]

    n_boletins        = len(historico)
    ultima_linha_hist = LINHA_HIST_INICIO + n_boletins - 1

    # ── H28: MEDIÇÕES BRUTAS = soma dinâmica do bloco ─────────
    if n_boletins > 0:
        _w(ws, "H28",
           f"=SUM(H{LINHA_HIST_INICIO}:H{ultima_linha_hist})")
    else:
        _w(ws, "H28", 0)

    # ── Linhas calculadas: posição FIXA confirmada no modelo ──
    # Essas linhas são fixas porque os textos das seções 5, 6 e 7
    # sempre ocupam o mesmo número de linhas no modelo padrão.
    # Se o número de boletins mudar, apenas o bloco D/H30:HN muda;
    # as seções abaixo (5, 6, 7) ficam no mesmo lugar no modelo base.

    _w(ws, "H57", valor_mes)
    _w(ws, "H59", "=H17-H28")
    _w(ws, "H61", "=H17-H28")
    _w(ws, "H63", valor_mes)

    # ══════════════════════════════════════════════════════════
    # ABA BOLETIM
    # ══════════════════════════════════════════════════════════
    ws_b = wb[NOME_BOLETIM]

    # ── Linha 1: título ───────────────────────────────────────
    # A1 é a célula superior esquerda da mesclagem A1:M1
    _w(ws_b, "A1",
       f"BOLETIM DE MEDIÇÃO Nº {num_medicao:02d} - {mes_execucao}")

    # ── Linha 2: empresa / contrato / modalidade ──────────────
    _w(ws_b, "E2", empresa) # E2 é a célula superior esquerda da mesclagem E2:G2
    _w(ws_b, "I2", contrato) # I2 é a célula superior esquerda da mesclagem I2:K2
    _w(ws_b, "M2", modalidade) # M2 é a célula superior esquerda da mesclagem M2:M2 (não mesclada, mas segue o padrão)

    # ── Linha 3: local / período ──────────────────────────────
    _w(ws_b, "E3", local) # E3 é a célula superior esquerda da mesclagem E3:G3
    _w(ws_b, "I3", periodo) # I3 é a célula superior esquerda da mesclagem I3:M3

    # ── Linha 4: data base ────────────────────────────────────
    _w(ws_b, "E4", _fmt_date(data_base)) # E4 é a célula superior esquerda da mesclagem E4:G4

    # ── Linha 7: dados do item ────────────────────────────────
    # (linhas 5 e 6 são cabeçalhos fixos da tabela)
    LINHA_ITEM = 7
    _wn(ws_b, LINHA_ITEM, 1,  item_num)            # A  ITEM
    _wn(ws_b, LINHA_ITEM, 2,  descricao_servico)   # B  DESCRIÇÃO (B7:C7 mesclada)
    _wn(ws_b, LINHA_ITEM, 4,  und)                 # D  UND
    _wn(ws_b, LINHA_ITEM, 5,  quant_total)         # E  QUANT. contrato
    _wn(ws_b, LINHA_ITEM, 6,  preco_unitario)      # F  P.U.
    _wn(ws_b, LINHA_ITEM, 7,  valor_original)      # G  PREVISTO
    _wn(ws_b, LINHA_ITEM, 8,  quant_acum_ant)      # H  ACUM.ANT. qtd
    _wn(ws_b, LINHA_ITEM, 9,  quant_mes)           # I  MES qtd
    _wn(ws_b, LINHA_ITEM, 10, quant_acum_total)    # J  ACUM.TOTAL qtd
    _wn(ws_b, LINHA_ITEM, 11, valor_acum_ant)      # K  ACUM.ANT. valor
    _wn(ws_b, LINHA_ITEM, 12, valor_mes)           # L  MES valor
    _wn(ws_b, LINHA_ITEM, 13, valor_acum_total)    # M  ACUM.TOTAL valor

    # ── Linha 28: TOTAL DO VALOR DO CONTRATO ─────────────────
    _wn(ws_b, 28, 7,  valor_original)   # G
    _wn(ws_b, 28, 11, valor_acum_ant)   # K
    _wn(ws_b, 28, 12, valor_mes)        # L
    _wn(ws_b, 28, 13, valor_acum_total) # M

    # ── Linha 30: MEDIÇÃO BRUTA ───────────────────────────────
    _wn(ws_b, 30, 11, valor_acum_ant)
    _wn(ws_b, 30, 12, valor_mes)
    _wn(ws_b, 30, 13, valor_acum_total)

    # ── Linha 31: FATURAMENTO DIRETO → 0 ─────────────────────
    _wn(ws_b, 31, 11, 0)
    _wn(ws_b, 31, 12, 0)
    _wn(ws_b, 31, 13, 0)

    # ── Linha 32: ADIANTAMENTO → 0 ───────────────────────────
    _wn(ws_b, 32, 11, 0)
    _wn(ws_b, 32, 12, 0)
    _wn(ws_b, 32, 13, 0)

    # ── Linha 33: MEDIÇÃO LÍQUIDA = BRUTA (sem descontos) ────
    _wn(ws_b, 33, 11, valor_acum_ant)
    _wn(ws_b, 33, 12, valor_mes)
    _wn(ws_b, 33, 13, valor_acum_total)

    wb.save(output_path)
    return output_path
