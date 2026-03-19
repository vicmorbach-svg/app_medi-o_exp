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
    Extrai mês por extenso + ano a partir do período.
    Aceita: "01/02/2026 A 28/02/2026" ou "01/02 A 28/02/2026"
    Retorna: "FEVEREIRO 2026"
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


def _w(ws, cell: str, value):
    """Escreve valor na célula preservando formatação existente."""
    ws[cell] = value


def gerar_excel_medicao(dados: dict) -> Path:
    """
    Copia o modelo e preenche exatamente as células mapeadas.
    Retorna o Path do arquivo gerado.
    """

    # ── Desempacota dados ────────────────────────────────────────────────────
    contrato          = dados["contrato"]
    num_medicao       = int(dados["num_medicao"])
    periodo           = dados["periodo"]
    data_apresentacao = dados["data_apresentacao"]
    descricao_servico = dados["descricao_servico"]
    valor_mes         = float(dados["valor_mes"])
    quant_mes         = float(dados["quant_mes"])

    empresa           = dados["empresa"]
    local             = dados["local"]
    modalidade        = dados["modalidade"]
    data_base         = dados["data_base"]
    data_termino      = dados["data_termino"]
    valor_original    = float(dados["valor_original"])
    item_num          = dados["item_num"]
    und               = dados["und"]
    quant_total       = float(dados["quant_total"])
    preco_unitario    = float(dados["preco_unitario"])
    centro_custo      = dados.get("centro_custo", "")
    conta_contabil    = dados.get("conta_contabil", "")
    item_caixa        = dados.get("item_caixa", "")

    quant_acum_ant    = float(dados["quant_acum_ant"])
    valor_acum_ant    = float(dados["valor_acum_ant"])
    quant_acum_total  = float(dados["quant_acum_total"])
    valor_acum_total  = float(dados["valor_acum_total"])
    saldo_contrato    = float(dados["saldo_contrato"])

    # Lista de boletins para o bloco amarelo
    # Cada item: {"label": str, "valor": float}
    historico         = dados.get("historico", [])

    # Mês extraído do período
    mes_execucao = extrair_mes_do_periodo(periodo)

    # Helper para formatar datas
    def fmt_date(d):
        if isinstance(d, date):
            return d.strftime("%d/%m/%Y")
        return str(d)

    # ── Nome do arquivo de saída ─────────────────────────────────────────────
    mes_slug = mes_execucao.replace(" ", "_")
    nome = f"medicao_{num_medicao:02d}_{contrato}_{mes_slug}.xlsx"
    output_path = OUTPUT_DIR / nome

    shutil.copy(MODELO_FILE, output_path)
    wb = openpyxl.load_workbook(output_path)

    # ════════════════════════════════════════════════════════════════════════
    # ABA PROTOCOLO
    # ════════════════════════════════════════════════════════════════════════
    ws = wb["PROTOCOLO"]

    # Linha 6: número do boletim
    _w(ws, "A6", f"Boletim de Medição n. {num_medicao:02d}")

    # Linha 7: descrição do serviço (selecionada pelo usuário)
    _w(ws, "A7", descricao_servico)

    # Dados fixos do contrato
    _w(ws, "B8",  empresa)
    _w(ws, "B9",  contrato)
    _w(ws, "H9",  modalidade)
    _w(ws, "D3",  contrato)
    _w(ws, "B10", local)
    _w(ws, "B11", fmt_date(data_base))
    _w(ws, "H12", fmt_date(data_termino))
    _w(ws, "H17", valor_original)
    _w(ws, "H18", valor_original)
    _w(ws, "H23", valor_original)
    _w(ws, "H25", valor_original)

    # Dados variáveis da medição
    _w(ws, "B12", periodo)
    _w(ws, "C14", mes_execucao)
    _w(ws, "C15", fmt_date(data_apresentacao))

    # ── Bloco amarelo: histórico de boletins ─────────────────────────────────
    LINHA_HIST_INICIO = 30  # linha real no Excel onde começa o histórico

    # Limpa bloco inteiro antes de reescrever (previne lixo de execuções anteriores)
    for r in range(LINHA_HIST_INICIO, LINHA_HIST_INICIO + 60):
        ws.cell(row=r, column=4, value=None)  # coluna D
        ws.cell(row=r, column=8, value=None)  # coluna H

    # Escreve cada boletim do histórico
    for i, h in enumerate(historico):
        row = LINHA_HIST_INICIO + i
        ws.cell(row=row, column=4, value=h["label"])
        ws.cell(row=row, column=8, value=h["valor"])

    # Linha final do bloco amarelo
    n_boletins = len(historico)
    ultima_linha_hist = LINHA_HIST_INICIO + n_boletins - 1

    # ── H28: MEDIÇÕES BRUTAS = soma dinâmica do bloco amarelo ────────────────
    if n_boletins > 0:
        _w(ws, "H28", f"=SUM(H{LINHA_HIST_INICIO}:H{ultima_linha_hist})")
    else:
        _w(ws, "H28", 0)

    # ── Linhas calculadas após o bloco amarelo ───────────────────────────────
    # A estrutura após os boletins é sempre:
    # +1 linha vazia
    # +2 "5.- DESCONTO DE MATERIAIS" → não mexemos
    # +3 a +6 linhas vazias/fixas → não mexemos
    # +7 "6.- RETENÇÕES" → não mexemos
    # +8 cabeçalho retenções → não mexemos
    # +9 linha vazia → não mexemos
    # +10 "7.- OBSERVAÇÃO" → não mexemos
    # +11 linha vazia → não mexemos
    # +12 linha vazia → não mexemos
    # +13 VALOR BRUTO
    # +14 linha vazia
    # +15 SALDO CONTRATO JURÍDICO
    # +16 linha vazia
    # +17 SALDO SOURCING
    # +18 linha vazia
    # +19 CENTRO DE CUSTO

    # No modelo com 15 boletins (linhas 30–44), ultima_linha_hist = 44
    # VALOR BRUTO estava na linha 57 → 44 + 13 = 57 ✓
    # Calculamos dinamicamente a partir de ultima_linha_hist:

    linha_valor_bruto    = ultima_linha_hist + 13
    linha_saldo_juridico = ultima_linha_hist + 15
    linha_saldo_sourcing = ultima_linha_hist + 17
    linha_centro_custo   = ultima_linha_hist + 19

    _w(ws, f"H{linha_valor_bruto}",    valor_mes)
    _w(ws, f"H{linha_saldo_juridico}", "=H17-H28")
    _w(ws, f"H{linha_saldo_sourcing}", "=H17-H28")

    # Linha do centro de custo: atualiza textos + valor
    _w(ws, f"A{linha_centro_custo}", f"CENTRO DE CUSTO: {centro_custo}")
    _w(ws, f"C{linha_centro_custo}", f"CONTA CONTÁBIL: {conta_contabil}")
    _w(ws, f"F{linha_centro_custo}", f"ITEM CAIXA: {item_caixa}")
    _w(ws, f"H{linha_centro_custo}", valor_mes)

    # ════════════════════════════════════════════════════════════════════════
    # ABA BOLETIM
    # ════════════════════════════════════════════════════════════════════════
    ws_b = wb["BOLETIM"]

    # Linha 1: título
    _w(ws_b, "A1", f"BOLETIM DE MEDIÇÃO Nº {num_medicao:02d} - {mes_execucao}")

    # Linha 2: empresa / contrato / modalidade
    _w(ws_b, "E2", empresa)
    _w(ws_b, "I2", contrato)
    _w(ws_b, "M2", modalidade)

    # Linha 3: local / período
    _w(ws_b, "E3", local)
    _w(ws_b, "I3", periodo)

    # Linha 4: data base
    _w(ws_b, "E4", fmt_date(data_base))

    # Linha 6: dados do item
    ws_b.cell(row=6, column=1,  value=item_num)           # A  ITEM
    ws_b.cell(row=6, column=2,  value=descricao_servico)  # B  DESCRIÇÃO
    ws_b.cell(row=6, column=4,  value=und)                # D  UND
    ws_b.cell(row=6, column=5,  value=quant_total)        # E  QUANT. contrato
    ws_b.cell(row=6, column=6,  value=preco_unitario)     # F  P.U.
    ws_b.cell(row=6, column=7,  value=valor_original)     # G  PREVISTO
    ws_b.cell(row=6, column=8,  value=quant_acum_ant)     # H  ACUM.ANT. qtd
    ws_b.cell(row=6, column=9,  value=quant_mes)          # I  MES qtd
    ws_b.cell(row=6, column=10, value=quant_acum_total)   # J  ACUM.TOTAL qtd
    ws_b.cell(row=6, column=11, value=valor_acum_ant)     # K  ACUM.ANT. valor
    ws_b.cell(row=6, column=12, value=valor_mes)          # L  MES valor
    ws_b.cell(row=6, column=13, value=valor_acum_total)   # M  ACUM.TOTAL valor

    # Linha 32: TOTAL DO VALOR DO CONTRATO
    ws_b.cell(row=32, column=7,  value=valor_original)
    ws_b.cell(row=32, column=11, value=valor_acum_ant)
    ws_b.cell(row=32, column=12, value=valor_mes)
    ws_b.cell(row=32, column=13, value=valor_acum_total)

    # Linha 34: MEDIÇÃO BRUTA
    ws_b.cell(row=34, column=11, value=valor_acum_ant)
    ws_b.cell(row=34, column=12, value=valor_mes)
    ws_b.cell(row=34, column=13, value=valor_acum_total)

    # Linhas 35 e 36: FATURAMENTO DIRETO e ADIANTAMENTO → 0
    for row in (35, 36):
        for col in (11, 12, 13):
            ws_b.cell(row=row, column=col, value=0)

    # Linha 37: MEDIÇÃO LÍQUIDA = mesma que medição bruta (sem descontos)
    ws_b.cell(row=37, column=11, value=valor_acum_ant)
    ws_b.cell(row=37, column=12, value=valor_mes)
    ws_b.cell(row=37, column=13, value=valor_acum_total)

    wb.save(output_path)
    return output_path
