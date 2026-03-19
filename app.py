import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import date, datetime

from config import CONTRATOS_FILE, MEDICOES_FILE
from excel_writer import gerar_excel_medicao, extrair_mes_do_periodo
from pdf_converter import gerar_pdfs_medicao

st.set_page_config(
    page_title="Medição de Serviços",
    page_icon="📋",
    layout="wide"
)

# ──────────────────────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────────────────────

def fmt_brl(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def to_date(val):
    if isinstance(val, date):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        return pd.to_datetime(val).date()
    except Exception:
        return None

def carregar_contratos() -> pd.DataFrame:
    if CONTRATOS_FILE.exists():
        return pd.read_excel(CONTRATOS_FILE)
    df = pd.DataFrame([{
        "contrato":             "4600072112",
        "empresa":              "4C DIGITAL",
        "local":                "Rua da Assembléia, nº 10, sala 3318, Rio de Janeiro",
        "modalidade":           "Unitário",
        "data_base":            "2024-08-29",
        "data_termino":         "2025-08-29",
        "valor_original":       1864800.00,
        "item_num":             20,
        "und":                  "un",
        "quant_total":          1864800,
        "preco_unitario":       1.00,
        "centro_custo":         "GBC1131003",
        "conta_contabil":       "",
        "item_caixa":           "",
        "servicos_disponiveis": "ENVIO DE SMS;SV DE COBRANÇA",
    }])
    df.to_excel(CONTRATOS_FILE, index=False)
    return df

def carregar_medicoes() -> pd.DataFrame:
    if MEDICOES_FILE.exists():
        return pd.read_excel(MEDICOES_FILE)
    df = pd.DataFrame(columns=[
        "contrato", "num_medicao", "data_apresentacao", "periodo",
        "mes_execucao", "descricao_servico",
        "quant_mes", "valor_mes",
        "quant_acum_ant", "valor_acum_ant",
        "quant_acum_total", "valor_acum_total",
        "valor_original", "saldo_contrato"
    ])
    df.to_excel(MEDICOES_FILE, index=False)
    return df

def salvar_medicoes(df: pd.DataFrame):
    df.to_excel(MEDICOES_FILE, index=False)

def calcular_acumulado(df: pd.DataFrame, contrato: str, num_medicao: int):
    ant = df[
        (df["contrato"].astype(str) == str(contrato)) &
        (df["num_medicao"].astype(int) < num_medicao)
    ]
    if ant.empty:
        return 0.0, 0.0
    ultima = ant.sort_values("num_medicao").iloc[-1]
    return float(ultima["quant_acum_total"]), float(ultima["valor_acum_total"])

def montar_historico(df: pd.DataFrame, contrato: str,
                     num_medicao_atual: int, valor_atual: float) -> list:
    """Retorna lista de todos os boletins (anteriores + atual) para o bloco amarelo."""
    anteriores = df[
        (df["contrato"].astype(str) == str(contrato)) &
        (df["num_medicao"].astype(int) < num_medicao_atual)
    ].sort_values("num_medicao")

    hist = []
    for _, r in anteriores.iterrows():
        hist.append({
            "label": f"Boletim de Medição n. {int(r['num_medicao']):02d} SMS",
            "valor": float(r["valor_mes"])
        })
    hist.append({
        "label": f"Boletim de Medição n. {num_medicao_atual:02d} SMS",
        "valor": valor_atual
    })
    return hist

# ──────────────────────────────────────────────────────────────────────────────
# ESTADO DA SESSÃO
# ──────────────────────────────────────────────────────────────────────────────

for key, default in [
    ("excel_path", None),
    ("pdf_paths", {}),
    ("dados_gerados", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ──────────────────────────────────────────────────────────────────────────────
# CARREGA DADOS
# ──────────────────────────────────────────────────────────────────────────────

df_contratos = carregar_contratos()
df_medicoes  = carregar_medicoes()

# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.header("📊 Histórico de Medições")
    if not df_medicoes.empty:
        st.dataframe(
            df_medicoes[[
                "contrato", "num_medicao", "periodo",
                "valor_mes", "valor_acum_total", "saldo_contrato"
            ]].sort_values(["contrato", "num_medicao"]),
            use_container_width=True
        )
    else:
        st.info("Nenhuma medição registrada ainda.")

    st.divider()
    if st.button("🔄 Recarregar dados"):
        st.rerun()

# ──────────────────────────────────────────────────────────────────────────────
# TÍTULO
# ──────────────────────────────────────────────────────────────────────────────

st.title("📋 Medição de Serviços")
st.divider()

# ──────────────────────────────────────────────────────────────────────────────
# 1. CONTRATO
# ──────────────────────────────────────────────────────────────────────────────

st.subheader("1. Contrato")

contrato_sel = st.selectbox(
    "Selecione o contrato",
    df_contratos["contrato"].astype(str).unique()
)
cr = df_contratos[
    df_contratos["contrato"].astype(str) == contrato_sel
].iloc[0]

c1, c2, c3, c4 = st.columns(4)
c1.metric("Empresa",        cr["empresa"])
c2.metric("Modalidade",     cr["modalidade"])
c3.metric("Valor Original", fmt_brl(float(cr["valor_original"])))
c4.metric("Centro Custo",   cr["centro_custo"])

with st.expander("Ver todos os dados do contrato"):
    col1, col2 = st.columns(2)
    col1.write(f"**Local:** {cr['local']}")
    col1.write(f"**Data Base:** {cr['data_base']}")
    col1.write(f"**Término:** {cr['data_termino']}")
    col2.write(f"**Item:** {cr['item_num']} — {cr['und']}")
    col2.write(f"**Quant. Total:** {int(cr['quant_total']):,}".replace(",", "."))
    col2.write(f"**P.U.:** {fmt_brl(float(cr['preco_unitario']))}")
    col2.write(f"**Conta Contábil:** {cr['conta_contabil']}")

st.divider()

# ──────────────────────────────────────────────────────────────────────────────
# 2. DADOS DA MEDIÇÃO
# ──────────────────────────────────────────────────────────────────────────────

st.subheader("2. Dados da Medição")

col1, col2, col3 = st.columns([1, 2, 2])

with col1:
    num_medicao = st.number_input("Nº do Boletim", min_value=1, step=1, value=1)

with col2:
    periodo = st.text_input(
        "Período",
        placeholder="01/02/2026 A 28/02/2026",
        help="Formato: DD/MM/AAAA A DD/MM/AAAA"
    )

with col3:
    data_apresentacao = st.date_input("Data de apresentação", value=date.today())

# Mês extraído automaticamente
mes_execucao = extrair_mes_do_periodo(periodo) if periodo else ""

if mes_execucao:
    st.success(f"Mês de execução: **{mes_execucao}**")
elif periodo:
    st.warning("Não foi possível extrair o mês. Verifique o formato do período.")

st.divider()

# ──────────────────────────────────────────────────────────────────────────────
# 3. SERVIÇO E QUANTIDADE
# ──────────────────────────────────────────────────────────────────────────────

st.subheader("3. Serviço e Quantidade")

servicos_lista = [
    s.strip()
    for s in str(cr.get("servicos_disponiveis", "")).split(";")
    if s.strip()
] or ["ENVIO DE SMS"]

col1, col2 = st.columns(2)

with col1:
    descricao_servico = st.selectbox("Descrição do serviço (A7)", servicos_lista)

with col2:
    quant_mes = st.number_input(
        f"Quantidade medida no mês ({cr['und']})",
        min_value=0.0, step=0.01, value=0.0, format="%.2f"
    )

# ── Cálculos ─────────────────────────────────────────────────────────────────

valor_original   = float(cr["valor_original"])
preco_unitario   = float(cr["preco_unitario"])

quant_acum_ant, valor_acum_ant = calcular_acumulado(
    df_medicoes, contrato_sel, num_medicao
)

valor_mes        = quant_mes * preco_unitario
quant_acum_total = quant_acum_ant + quant_mes
valor_acum_total = valor_acum_ant + valor_mes
saldo_contrato   = valor_original - valor_acum_total

st.divider()

# ──────────────────────────────────────────────────────────────────────────────
# 4. RESUMO
# ──────────────────────────────────────────────────────────────────────────────

st.subheader("4. Resumo da Medição")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Acum. Anterior",   fmt_brl(valor_acum_ant))
c2.metric("Valor do Mês",     fmt_brl(valor_mes))
c3.metric("Acum. Total",      fmt_brl(valor_acum_total))
c4.metric("Saldo do Contrato", fmt_brl(saldo_contrato))

pct = (valor_acum_total / valor_original * 100) if valor_original > 0 else 0
st.progress(
    min(pct / 100, 1.0),
    text=f"Execução contratual: {pct:.1f}%"
)

st.divider()

# ──────────────────────────────────────────────────────────────────────────────
# 5. GERAR ARQUIVOS
# ──────────────────────────────────────────────────────────────────────────────

st.subheader("5. Gerar Arquivos")

campos_ok = bool(periodo and mes_execucao and quant_mes > 0)

if not campos_ok:
    st.info("Preencha o período e a quantidade para habilitar a geração.")

if st.button("⚙️ Gerar Excel + PDF", type="primary",
             use_container_width=True, disabled=not campos_ok):

    historico = montar_historico(
        df_medicoes, contrato_sel, num_medicao, valor_mes
    )

    dados_medicao = {
        "contrato":           contrato_sel,
        "num_medicao":        num_medicao,
        "periodo":            periodo,
        "data_apresentacao":  data_apresentacao,
        "descricao_servico":  descricao_servico,
        "valor_mes":          valor_mes,
        "quant_mes":          quant_mes,
        "empresa":            cr["empresa"],
        "local":              cr["local"],
        "modalidade":         cr["modalidade"],
        "data_base":          to_date(cr["data_base"]),
        "data_termino":       to_date(cr["data_termino"]),
        "valor_original":     valor_original,
        "item_num":           cr["item_num"],
        "und":                cr["und"],
        "quant_total":        float(cr["quant_total"]),
        "preco_unitario":     preco_unitario,
        "centro_custo":       str(cr.get("centro_custo", "")),
        "conta_contabil":     str(cr.get("conta_contabil", "")),
        "item_caixa":         str(cr.get("item_caixa", "")),
        "quant_acum_ant":     quant_acum_ant,
        "valor_acum_ant":     valor_acum_ant,
        "quant_acum_total":   quant_acum_total,
        "valor_acum_total":   valor_acum_total,
        "saldo_contrato":     saldo_contrato,
        "historico":          historico,
    }

    with st.spinner("Gerando Excel..."):
        try:
            excel_path = gerar_excel_medicao(dados_medicao)
            st.session_state.excel_path   = excel_path
            st.session_state.dados_gerados = dados_medicao
            st.success(f"Excel gerado: `{excel_path.name}`")
        except Exception as e:
            st.error(f"Erro ao gerar Excel: {e}")
            st.stop()

    with st.spinner("Convertendo para PDF..."):
        try:
            pdfs = gerar_pdfs_medicao(excel_path)
            st.session_state.pdf_paths = pdfs
            st.success("PDFs gerados!")
        except Exception as e:
            st.warning(f"PDF não gerado automaticamente: {e}")
            st.session_state.pdf_paths = {}

# ──────────────────────────────────────────────────────────────────────────────
# 6. DOWNLOADS
# ──────────────────────────────────────────────────────────────────────────────

if st.session_state.excel_path and Path(st.session_state.excel_path).exists():
    st.divider()
    st.subheader("6. Downloads")

    c1, c2, c3 = st.columns(3)

    c1.download_button(
        "📥 Baixar Excel",
        data=Path(st.session_state.excel_path).read_bytes(),
        file_name=Path(st.session_state.excel_path).name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    for aba, col in [("PROTOCOLO", c2), ("BOLETIM", c3)]:
        if aba in st.session_state.pdf_paths:
            p = Path(st.session_state.pdf_paths[aba])
            if p.exists():
                col.download_button(
                    f"📄 PDF — {aba}",
                    data=p.read_bytes(),
                    file_name=p.name,
                    mime="application/pdf",
                    use_container_width=True
                )

# ──────────────────────────────────────────────────────────────────────────────
# 7. SALVAR MEDIÇÃO
# ──────────────────────────────────────────────────────────────────────────────

    st.divider()
    st.subheader("7. Confirmar e Salvar")

    st.warning(
        "⚠️ Confira o Excel antes de salvar. "
        "Esta ação registra a medição no histórico."
    )

    if st.button("💾 Salvar medição", use_container_width=True, type="primary"):
        d = st.session_state.dados_gerados
        if d:
            # Remove duplicata se existir
            mask = ~(
                (df_medicoes["contrato"].astype(str) == str(contrato_sel)) &
                (df_medicoes["num_medicao"].astype(int) == int(num_medicao))
            )
            df_upd = df_medicoes[mask]

            novo = {
                "contrato":          d["contrato"],
                "num_medicao":       d["num_medicao"],
                "data_apresentacao": str(d["data_apresentacao"]),
                "periodo":           d["periodo"],
                "mes_execucao":      mes_execucao,
                "descricao_servico": d["descricao_servico"],
                "quant_mes":         d["quant_mes"],
                "valor_mes":         d["valor_mes"],
                "quant_acum_ant":    d["quant_acum_ant"],
                "valor_acum_ant":    d["valor_acum_ant"],
                "quant_acum_total":  d["quant_acum_total"],
                "valor_acum_total":  d["valor_acum_total"],
                "valor_original":    d["valor_original"],
                "saldo_contrato":    d["saldo_contrato"],
            }

            df_upd = pd.concat(
                [df_upd, pd.DataFrame([novo])],
                ignore_index=True
            )
            salvar_medicoes(df_upd)
            st.success("Medição salva com sucesso!")
            st.rerun()
