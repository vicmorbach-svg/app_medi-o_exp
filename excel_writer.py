import shutil
import openpyxl
from pathlib import Path
from datetime import date
from config import MODELO_FILE, OUTPUT_DIR

MESES_PT = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
}

def extrair_mes_do_periodo(periodo: str) -> str:
    """
    Extrai o mês por extenso do período informado.
    Exemplo: "01/02/2026 A 28/02/2026" → "FEVEREIRO 2026"
             "01/02 A 28/02/2026"      → tenta extrair do final
    """
    try:
        partes = periodo.upper().strip().split("A")
        data_fim = partes[-1].strip()
        # Tenta extrair DD/MM/YYYY ou DD/MM/AAAA
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


def _set(ws, cell: str, value):
    """Define valor em célula, preservando formatação existente."""
    ws[cell] = value


def gerar_excel_medicao(dados: dict) -> Path:
    """
    Copia o modelo e preenche exatamente as células mapeadas.
    Retorna o Path do arquivo gerado.
    """

    # ── Extrai dados ────────────────────────────────────────────────────────
    contrato            = dados["contrato"]
    num_medicao         = dados["num_medicao"]
    periodo             = dados["periodo"]
    data_apresentacao   = dados["data_apresentacao"]   # date
    descricao_servico   = dados["descricao_servico"]
    valor_mes           = dados["valor_mes"]
    quant_mes           = dados["quant_mes"]

    empresa             = dados["empresa"]
    local               = dados["local"]
    modalidade          = dados["modalidade"]
    data_base           = dados["data_base"]           # date
    data_termino        = dados["data_termino"]        # date
    valor_original      = dados["valor_original"]
    item_num            = dados["item_num"]
    und                 = dados["und"]
    quant_total         = dados["quant_total"]
    preco_unitario      = dados["preco_unitario"]

    quant_acum_ant      = dados["quant_acum_ant"]
    valor_acum_ant      = dados["valor_acum_ant"]
    quant_acum_total    = dados["quant_acum_total"]
    valor_acum_total    = dados["valor_acum_total"]
    saldo_contrato      = dados["saldo_contrato"]

    # Lista de medições anteriores: [{"label": str, "valor": float}, ...]
    historico           = dados.get("historico", [])

    # Mês de execução extraído automaticamente do período
    mes_execucao = extrair_mes_do_periodo(periodo)

    # ── Nome do arquivo de saída ─────────────────────────────────────────────
    mes_slug = mes_execucao.replace(" ", "_")
    nome_arquivo = f"medicao_{num_medicao:02d}_{contrato}_{mes_slug}.xlsx"
    output_path = OUTPUT_DIR / nome_arquivo

    shutil.copy(MODELO_FILE, output_path)
    wb = openpyxl.load_workbook(output_path)

    # ════════════════════════════════════════════════════════════════════════
    # ABA PROTOCOLO
    # ════════════════════════════════════════════════════════════════════════
    ws = wb["PROTOCOLO"]

    # ── Dados variáveis ──────────────────────────────────────────────────────
    _set(ws, "A6", f"Boletim de Medição n. {num_medicao:02d}")
    _set(ws, "A7", descricao_servico)
    _set(ws, "B12", periodo)
    _set(ws, "C14", mes_execucao)
    _set(ws, "C15", data_apresentacao.strftime("%d/%m/%Y") if isinstance(data_apresentacao, date) else str(data_apresentacao))

    # ── Dados fixos do contrato (verde) ──────────────────────────────────────
    _set(ws, "B8",  empresa)
    _set(ws, "B9",  contrato)
    _set(ws, "D3",  contrato)
    _set(ws, "B10", local)
    _set(ws, "B11", data_base.strftime("%d/%m/%Y") if isinstance(data_base, date) else str(data_base))
    _set(ws, "H12", data_termino.strftime("%d/%m/%Y") if isinstance(data_termino, date) else str(data_termino))
    _set(ws, "H17", valor_original)
    _set(ws, "H18", valor_original)
    _set(ws, "H23", valor_original)
    _set(ws, "H25", valor_original)

    # ── Bloco cumulativo amarelo (histórico de boletins) ─────────────────────
    # Começa na linha 30, coluna D (texto) e H (valor)
    LINHA_HIST_INICIO = 30

    # Limpa linhas do bloco (até 60 medições previstas)
    for r in range(LINHA_HIST_INICIO, LINHA_HIST_INICIO + 60):
        ws.cell(row=r, column=4, value=None)   # D
        ws.cell(row=r, column=8, value=None)   # H

    # Escreve histórico linha a linha
    for i, h in enumerate(historico):
        row = LINHA_HIST_INICIO + i
        ws.cell(row=row, column=4, value=h["label"])
        ws.cell(row=row, column=8, value=h["valor"])

    # Linha final do bloco = última linha do histórico
    ultima_linha_hist = LINHA_HIST_INICIO + len(historico) - 1

    # ── H28: MEDIÇÕES BRUTAS = soma do bloco cumulativo ──────────────────────
    # Escrevemos como fórmula dinâmica
    if historico:
        _set(ws, "H28", f"=SUM(H{LINHA_HIST_INICIO}:H{ultima_linha_hist})")
    else:
        _set(ws, "H28", 0)

    # ── Células calculadas ───────────────────────────────────────────────────
    # H57: Valor bruto da medição no mês
    _set(ws, "H57", valor_mes)

    # H59: Saldo do Contrato Jurídico = H17 - H28
    _set(ws, "H59", "=H17-H28")

    # H61: Saldo da Sourcing = H17 - H28
    _set(ws, "H61", "=H17-H28")

    # H63: linha de CENTRO DE CUSTO / CONTA CONTÁBIL / ITEM CAIXA
    _set(ws, "H63", valor_mes)

    # ════════════════════════════════════════════════════════════════════════
    # ABA BOLETIM
    # ════════════════════════════════════════════════════════════════════════
    ws_b = wb["BOLETIM"]

    # ── Título ───────────────────────────────────────────────────────────────
    _set(ws_b, "A1", f"BOLETIM DE MEDIÇÃO Nº {num_medicao:02d} - {mes_execucao}")

    # ── Cabeçalho fixo do contrato ───────────────────────────────────────────
    _set(ws_b, "E2", empresa)
    _set(ws_b, "I2", contrato)
    _set(ws_b, "M2", modalidade)
    _set(ws_b, "E3", local)
    _set(ws_b, "E4", data_base.strftime("%d/%m/%Y") if isinstance(data_base, date) else str(data_base))

    # ── Período (variável) ───────────────────────────────────────────────────
    _set(ws_b, "I3", periodo)

    # ── Linha do item (linha 6) ──────────────────────────────────────────────
    # Fixos do contrato
    ws_b.cell(row=6, column=1, value=item_num)          # A6  ITEM
    ws_b.cell(row=6, column=2, value=descricao_servico) # B6  DESCRIÇÃO
    ws_b.cell(row=6, column=4, value=und)               # D6  UND
    ws_b.cell(row=6, column=5, value=quant_total)       # E6  QUANT. total contrato
    ws_b.cell(row=6, column=6, value=preco_unitario)    # F6  P.U.
    ws_b.cell(row=6, column=7, value=valor_original)    # G6  PREVISTO

    # Calculados / variáveis
    ws_b.cell(row=6, column=8,  value=quant_acum_ant)   # H6  ACUM. ANT. qtd
    ws_b.cell(row=6, column=9,  value=quant_mes)        # I6  MES qtd
    ws_b.cell(row=6, column=10, value=quant_acum_total) # J6  ACUM. TOTAL qtd
    ws_b.cell(row=6, column=11, value=valor_acum_ant)   # K6  ACUM. ANT. valor
    ws_b.cell(row=6, column=12, value=valor_mes)        # L6  MES valor
    ws_b.cell(row=6, column=13, value=valor_acum_total) # M6  ACUM. TOTAL valor

    # ── Totais (linha 32) ────────────────────────────────────────────────────
    ws_b.cell(row=32, column=7,  value=valor_original)   # G32 total contrato
    ws_b.cell(row=32, column=11, value=valor_acum_ant)   # K32
    ws_b.cell(row=32, column=12, value=valor_mes)        # L32
    ws_b.cell(row=32, column=13, value=valor_acum_total) # M32

    # ── Medição Bruta (linha 34) ─────────────────────────────────────────────
    ws_b.cell(row=34, column=11, value=valor_acum_ant)
    ws_b.cell(row=34, column=12, value=valor_mes)
    ws_b.cell(row=34, column=13, value=valor_acum_total)

    # ── Faturamento Direto (linha 35) e Adiantamento (linha 36) → 0 ─────────
    for row in (35, 36):
        for col in (11, 12, 13):
            ws_b.cell(row=row, column=col, value=0)

    # ── Medição Líquida (linha 37) = Medição Bruta ───────────────────────────
    ws_b.cell(row=37, column=11, value=valor_acum_ant)
    ws_b.cell(row=37, column=12, value=valor_mes)
    ws_b.cell(row=37, column=13, value=valor_acum_total)

    wb.save(output_path)
    return output_path
