import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import date, datetime
import io

from config import CONTRATOS_FILE, MEDICOES_FILE, OUTPUT_DIR
from excel_writer import gerar_excel_medicao, extrair_mes_do_periodo
from pdf_converter import gerar_pdfs_medicao

try:
    from sharepoint import atualizar_acompanhamento, upload_arquivo_sharepoint
    SHAREPOINT_OK = True
except Exception:
    SHAREPOINT_OK = False

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
    if pd.isna(val):
        return None
    return pd.to_datetime(val).date()

def carregar_contratos() -> pd.DataFrame:
    if CONTRATOS_FILE.exists():
        return pd.read_excel(CONTRATOS_FILE)
    # Cria arquivo de exemplo
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
    cols = [
        "contrato", "num_medicao", "data_apresentacao", "periodo",
        "mes_execucao", "descricao_servico",
        "quant_mes", "valor_mes",
        "quant_acum_ant", "valor_acum_ant",
        "quant_acum_total", "valor_acum_total",
        "valor_original", "saldo_contrato"
    ]
    df = pd.DataFrame(columns=cols)
    df.to_excel(MEDICOES_FILE, index=False)
    return df

def salvar_medicoes(df: pd.DataFrame):
    df.to_excel(MEDICOES_FILE, index=False)

def calcular_acumulado(df: pd.DataFrame, contrato: str, num_medicao: int):
    """Retorna (quant_acum_ant, valor_acum_ant) das medições anteriores."""
    ant = df[(df["contrato"] == contrato) & (df["num_medicao"] < num_medicao)]
    if ant.empty:
        return 0.0, 0.0
    ultima = ant.sort_values("num_medicao").iloc[-1]
    return float(ultima["quant_acum_total"]), float(ultima["valor_acum_total"])

def montar_historico(df: pd.DataFrame, contrato: str, num_medicao_atual: int,
                     descricao_atual: str, valor_atual: float) -> list:
    """
    Monta lista de boletins para o bloco amarelo do PROTOCOLO.
    Inclui todos os anteriores + o atual.
    """
    anteriores = df[
        (df["contrato"] == contrato) &
        (df["num_medicao"] < num_medicao_atual)
    ].sort_values("num_medicao")

    hist = []
    for _, r in anteriores.iterrows():
        hist.append({
            "label": f"Boletim de Medição n. {int(r['num_medicao']):02d} SMS",
            "valor": float(r["valor_mes"])
        })

    # Adiciona medição atual
    hist.append({
        "label": f"Boletim de Medição n. {num_medicao_atual:02d} SMS",
        "valor": valor_atual
    })
    return hist

# ──────────────────────────────────────────────────────────────────────────────
# ESTADO DA SESSÃO
# ──────────────────────────────────────────────────────────────────────────────

for key in ["excel_path", "pdf_paths", "dados_gerados"]:
    if key not in st.session_state:
        st.session_state[key] = None if key != "pdf_paths" else {}

# ──────────────────────────────────────────────────────────────────────────────
# DADOS
# ──────────────────────────────────────────────────────────────────────────────

df_contratos = carregar_contratos()
df_medicoes  = carregar_medicoes()

# ──────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.header("📊 Histórico de Medições")
    if not df_medicoes.empty:
        cols_show = ["contrato", "num_medicao", "periodo", "descricao_servico",
                     "valor_mes", "valor_acum_total", "saldo_contrato"]
        cols_exist = [c for c in cols_show if c in df_medicoes.columns]
        st.dataframe(
            df_medicoes[cols_exist].sort_values(["contrato", "num_medicao"]),
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
# SEÇÃO 1 — CONTRATO
# ──────────────────────────────────────────────────────────────────────────────

st.subheader("1. Contrato")

contrato_sel = st.selectbox(
    "Selecione o contrato",
    df_contratos["contrato"].astype(str).unique()
)

cr = df_contratos[df_contratos["contrato"].astype(str) == contrato_sel].iloc[0]

# Dados do contrato em cards
c1, c2, c3, c4 = st.columns(4)
c1.metric("Empresa",         cr["empresa"])
c2.metric("Modalidade",      cr["modalidade"])
c3.metric("Valor Original",  fmt_brl(float(cr["valor_original"])))
c4.metric("Centro de Custo", cr["centro_custo"])

with st.expander("Ver todos os dados do contrato"):
    col1, col2 = st.columns(2)
    col1.write(f"**Local:** {cr['local']}")
    col1.write(f"**Data Base:** {cr['data_base']}")
    col1.write(f"**Término:** {cr['data_termino']}")
    col2.write(f"**Item:** {cr['item_num']} — {cr['und']}")
    col2.write(f"**Quant. Total:** {cr['quant_total']:,.0f}".replace(",", "."))
    col2.write(f"**P.U.:** {fmt_brl(float(cr['preco_unitario']))}")
    col2.write(f"**Conta Contábil:** {cr['conta_contabil']}")

st.divider()

# ──────────────────────────────────────────────────────────────────────────────
# SEÇÃO 2 — DADOS DA MEDIÇÃO
# ──────────────────────────────────────────────────────────────────────────────

st.subheader("2. Dados da Medição")

col1, col2, col3 = st.columns([1, 2, 2])

with col1:
    num_medicao = st.number_input(
        "Nº do Boletim", min_value=1, step=1, value=1
    )

with col2:
    periodo = st.text_input(
        "Período",
        placeholder="01/02/2026 A 28/02/2026",
        help="Formato: DD/MM/AAAA A DD/MM/AAAA"
    )

with col3:
    data_apresentacao = st.date_input(
        "Data de apresentação",
        value=date.today()
    )

# Mês extraído automaticamente
mes_execucao = ""
if periodo:
    mes_execucao = extrair_mes_do_periodo(periodo)

if mes_execucao:
    st.success(f"Mês de execução identificado automaticamente: **{mes_execucao}**")
else:
    if periodo:
        st.warning("Não foi possível extrair o mês do período. Verifique o formato.")

st.divider()

# ──────────────────────────────────────────────────────────────────────────────
# SEÇÃO 3 — SERVIÇO E QUANTIDADE
# ──────────────────────────────────────────────────────────────────────────────

st.subheader("3. Serviço e Quantidade")

# Lista de serviços disponíveis para o contrato
servicos_raw = str(cr.get("servicos_disponiveis", ""))
servicos_lista = [s.strip() for s in servicos_raw.split(";") if s.strip()]

if not servicos_lista:
    servicos_lista = ["ENVIO DE SMS"]

descricao_servico = st.selectbox(
    "Descrição do serviço (A7)",
    options=servicos_lista,
    help="Serviços cadastrados para este contrato em contratos.xlsx"
)

col1, col2 = st.columns(2)

with col1:
    quant_mes = st.number_input(
        f"Quantidade medida no mês ({cr['und']})",
        min_value=0.0,
        step=0.01,
        value=0.0,
        format="%.2f"
    )

# ── Cálculos automáticos ──────────────────────────────────────────────────────

valor_original   = float(cr["valor_original"])
preco_unitario   = float(cr["preco_unitario"])

quant_acum_ant, valor_acum_ant = calcular_acumulado(
    df_medicoes, contrato_sel, num_medicao
)

valor_mes        = quant_mes * preco_unitario
quant_acum_total = quant_acum_ant + quant_mes
valor_acum_total = valor_acum_ant + valor_mes
saldo_contrato   = valor_original - valor_acum_total

with col2:
    st.metric("Valor calculado do mês", fmt_brl(valor_mes))

st.divider()

# ──────────────────────────────────────────────────────────────────────────────
# SEÇÃO 4 — RESUMO
# ──────────────────────────────────────────────────────────────────────────────

st.subheader("4. Resumo da Medição")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Acum. Anterior",   fmt_brl(valor_acum_ant),
            f"Qtd: {quant_acum_ant:,.2f}".replace(",","X").replace(".",",").replace("X","."))
col2.metric("Valor do Mês",     fmt_brl(valor_mes),
            f"Qtd: {quant_mes:,.2f}".replace(",","X").replace(".",",").replace("X","."))
col3.metric("Acum. Total",      fmt_brl(valor_acum_total),
            f"Qtd: {quant_acum_total:,.2f}".replace(",","X").replace(".",",").replace("X","."))
col4.metric("Saldo do Contrato", fmt_brl(saldo_contrato),
            delta_color="inverse")

# Progresso de execução do contrato
pct = (valor_acum_total / valor_original * 100) if valor_original > 0 else 0
st.progress(min(pct / 100, 1.0), text=f"Execução contratual: {pct:.1f}%")

st.divider()

# ──────────────────────────────────────────────────────────────────────────────
# SEÇÃO 5 — GERAR ARQUIVOS
# ──────────────────────────────────────────────────────────────────────────────

st.subheader("5. Gerar Arquivos")

campos_ok = bool(periodo and mes_execucao and quant_mes > 0)

if not campos_ok:
    st.info("Preencha o período e a quantidade para gerar os arquivos.")

if st.button(
    "⚙️ Gerar Excel + PDF",
    type="primary",
    use_container_width=True,
    disabled=not campos_ok
):
    with st.spinner("Montando histórico de boletins..."):
        historico = montar_historico(
            df_medicoes, contrato_sel, num_medicao,
            descricao_servico, valor_mes
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
            st.session_state.excel_path  = excel_path
            st.session_state.dados_gerados = dados_medicao
            st.success(f"Excel gerado: `{excel_path.name}`")
        except Exception as e:
            st.error(f"Erro ao gerar Excel: {e}")
            st.stop()

    with st.spinner("Convertendo para PDF..."):
        try:
            pdfs = gerar_pdfs_medicao(excel_path)
            st.session_state.pdf_paths = pdfs
            st.success("PDFs gerados com sucesso!")
        except Exception as e:
            st.warning(f"PDF não pôde ser gerado automaticamente: {e}")
            st.session_state.pdf_paths = {}

# ──────────────────────────────────────────────────────────────────────────────
# SEÇÃO 6 — DOWNLOADS
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

    if "PROTOCOLO" in st.session_state.pdf_paths:
        p = Path(st.session_state.pdf_paths["PROTOCOLO"])
        if p.exists():
            c2.download_button(
                "📄 PDF — PROTOCOLO",
                data=p.read_bytes(),
                file_name=p.name,
                mime="application/pdf",
                use_container_width=True
            )

    if "BOLETIM" in st.session_state.pdf_paths:
        p = Path(st.session_state.pdf_paths["BOLETIM"])
        if p.exists():
            c3.download_button(
                "📄 PDF — BOLETIM",
                data=p.read_bytes(),
                file_name=p.name,
                mime="application/pdf",
                use_container_width=True
            )

# ──────────────────────────────────────────────────────────────────────────────
# SEÇÃO 7 — CONFIRMAR E SALVAR
# ──────────────────────────────────────────────────────────────────────────────

    st.divider()
    st.subheader("7. Confirmar e Salvar")

    st.warning(
        "⚠️ Clique em **Salvar** somente após conferir o Excel gerado. "
        "Esta ação registra a medição no histórico e não pode ser desfeita facilmente."
    )

    col_local, col_sp = st.columns(2)

    with col_local:
        if st.button("💾 Salvar medição localmente", use_container_width=True):
            d = st.session_state.dados_gerados
            if d:
                # Remove duplicata se existir
                mask = ~(
                    (df_medicoes["contrato"] == contrato_sel) &
                    (df_medicoes["num_medicao"] == num_medicao)
                )
                df_medicoes_upd = df_medicoes[mask]

                novo = {
                    "contrato":          d["contrato"],
                    "num_medicao":       d["num_medicao"],
                    "data_apresentacao": d["data_apresentacao"],
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
                df_medicoes_upd = pd.concat(
                    [df_medicoes_upd, pd.DataFrame([novo])],
                    ignore_index=True
                )
                salvar_medicoes(df_medicoes_upd)
                st.success("Medição salva em `dados/medicoes.xlsx`!")
                st.rerun()

    with col_sp:
        if SHAREPOINT_OK:
            if st.button("☁️ Enviar ao SharePoint", use_container_width=True):
                d = st.session_state.dados_gerados
                with st.spinner("Enviando arquivos ao SharePoint..."):
                    try:
                        mes_slug = mes_execucao.replace(" ", "_")
                        base_nome = f"medicao_{num_medicao:02d}_{contrato_sel}_{mes_slug}"

                        # Envia Excel
                        upload_arquivo_sharepoint(
                            Path(st.session_state.excel_path).read_bytes(),
                            f"/Medicoes/{contrato_sel}/{base_nome}.xlsx"
                        )

                        # Envia PDFs
                        for aba, pdf_p in st.session_state.pdf_paths.items():
                            if Path(pdf_p).exists():
                                upload_arquivo_sharepoint(
                                    Path(pdf_p).read_bytes(),
                                    f"/Medicoes/{contrato_sel}/{base_nome}_{aba}.pdf"
                                )

                        # Atualiza planilha de acompanhamento
                        atualizar_acompanhamento({
                            "Contrato":            d["contrato"],
                            "Empresa":             d["empresa"],
                            "Num_Medicao":         d["num_medicao"],
                            "Periodo":             d["periodo"],
                            "Mes_Execucao":        mes_execucao,
                            "Data_Apresentacao":   str(d["data_apresentacao"]),
                            "Descricao_Servico":   d["descricao_servico"],
                            "Valor_Contrato":      d["valor_original"],
                            "Valor_Bruto_Mes":     d["valor_mes"],
                            "Valor_Acum_Anterior": d["valor_acum_ant"],
                            "Valor_Acum_Total":    d["valor_acum_total"],
                            "Saldo_Contrato":      d["saldo_contrato"],
                        })

                        st.success("Tudo enviado ao SharePoint!")
                    except Exception as e:
                        st.error(f"Erro ao enviar ao SharePoint: {e}")
        else:
            st.info(
                "Integração SharePoint não configurada. "
                "Preencha as credenciais em `config.py`."
            )
