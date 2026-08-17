from io import BytesIO
from pathlib import Path
import os

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st


st.set_page_config(
    page_title="Projeto SESAP | Painel Executivo",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

ROOT = Path(__file__).resolve().parent
SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1B7OFqNh8Pd2K3UsX7MCEkUdNYFAeloKe3CgCCYA8TiU")
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
LOCAL_FILE = ROOT / "dados_projeto.xlsx"

CORES = {
    "azul": "#005A70",
    "verde_agua": "#00A6A6",
    "verde": "#238636",
    "amarelo": "#F2B134",
    "vermelho": "#C93C37",
    "texto": "#17313A",
    "cinza": "#60747B",
    "fundo": "#F4F7F8",
}

st.markdown(
    """
    <style>
    .stApp { background: #F4F7F8; color: #17313A; }
    [data-testid="stSidebar"] { background: #FFFFFF; border-right: 1px solid #DDE6E8; }
    [data-testid="stMetric"] { background: white; border: 1px solid #DDE6E8; border-radius: 12px; padding: 16px; }
    [data-testid="stMetricLabel"] { color: #60747B; }
    [data-testid="stMetricValue"] { color: #17313A; }
    .hero { padding: 24px 28px; border-radius: 16px; color: white; margin-bottom: 18px;
            background: linear-gradient(110deg, #004E61 0%, #007D86 60%, #00A6A6 100%); }
    .hero h1 { margin: 0; font-size: 2rem; font-weight: 700; }
    .hero p { margin: 8px 0 0; color: #E6F7F7; }
    .section-title { font-size: 1.15rem; font-weight: 700; color: #17313A; margin: 12px 0 8px; }
    .phase-card { background: white; border: 1px solid #DDE6E8; border-radius: 12px; padding: 14px 16px; min-height: 112px; }
    .phase-active { border: 2px solid #00A6A6; box-shadow: 0 3px 12px rgba(0,166,166,.12); }
    .phase-name { font-weight: 700; color: #17313A; }
    .phase-period { font-size: .86rem; color: #60747B; margin-top: 4px; }
    .phase-status { display:inline-block; margin-top:12px; padding:4px 9px; border-radius:99px; font-size:.78rem; font-weight:600; }
    .done { background:#E5F5E9; color:#238636; } .active { background:#DDF5F3; color:#006D72; }
    .future { background:#EEF1F2; color:#60747B; }
    .note { background:#FFF9E8; border-left:4px solid #F2B134; padding:12px 14px; border-radius:6px; color:#5C4A12; }
    footer { visibility: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)


def fase_por_data(data):
    if pd.isna(data):
        return "Sem data"
    if data < pd.Timestamp("2025-10-01"):
        return "Fase 1 - Estratégica"
    if data < pd.Timestamp("2026-10-01"):
        return "Fase 2 - Tática"
    if data <= pd.Timestamp("2027-10-31"):
        return "Fase 3 - Operacional"
    return "Fora do período"


@st.cache_data(ttl=600, show_spinner=False)
def carregar_dados():
    origem = "Google Sheets"
    try:
        resposta = requests.get(SHEET_URL, timeout=20)
        resposta.raise_for_status()
        arquivo = BytesIO(resposta.content)
        planilha = pd.ExcelFile(arquivo)
    except Exception:
        if not LOCAL_FILE.exists():
            raise
        origem = "cópia local"
        planilha = pd.ExcelFile(LOCAL_FILE)

    entregas = pd.read_excel(planilha, sheet_name="Base_Dashboard", usecols="A:K").dropna(how="all")
    entregas.columns = ["ID", "Data", "Local", "Tipo", "Subtipo", "Descrição", "Frente", "Responsável", "Horas", "Status", "Lista Mestra"]
    entregas["Data"] = pd.to_datetime(entregas["Data"], errors="coerce")
    entregas["Horas"] = pd.to_numeric(entregas["Horas"], errors="coerce")
    entregas["Fase"] = entregas["Data"].map(fase_por_data)
    entregas["Mês"] = entregas["Data"].dt.to_period("M").dt.to_timestamp()
    for coluna in ["Local", "Tipo", "Subtipo", "Frente", "Responsável", "Status", "Lista Mestra"]:
        entregas[coluna] = entregas[coluna].fillna("Não informado").astype(str).str.strip()

    partes = []
    nomes = {1: "Estratégica", 2: "Tática", 3: "Operacional"}
    for numero in (1, 2, 3):
        marcos = pd.read_excel(planilha, sheet_name=f"Evolução Marcos Fase {numero}", usecols="A:G")
        marcos = marcos[marcos.iloc[:, 0].notna() & marcos.iloc[:, 1].notna()].copy()
        marcos.columns = ["ID", "Marco", "Início", "Fim", "Duração", "Progresso", "Progresso Ponderado"]
        marcos["Fase"] = f"Fase {numero} - {nomes[numero]}"
        marcos["Início"] = pd.to_datetime(marcos["Início"], errors="coerce")
        marcos["Fim"] = pd.to_datetime(marcos["Fim"], errors="coerce")
        marcos["Progresso"] = pd.to_numeric(marcos["Progresso"], errors="coerce").fillna(0)
        marcos["Situação"] = marcos["Progresso"].apply(lambda x: "Concluído" if x >= 100 else ("Em andamento" if x > 0 else "Não iniciado"))
        partes.append(marcos)
    return entregas, pd.concat(partes, ignore_index=True), origem


def grafico_layout(fig, altura=360):
    fig.update_layout(
        height=altura, margin=dict(l=15, r=15, t=42, b=15),
        paper_bgcolor="white", plot_bgcolor="white", font=dict(color=CORES["texto"]),
        title_font=dict(size=16), legend_title_text="",
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#E8EEF0")
    return fig


def progresso_ponderado(df):
    duracao = pd.to_numeric(df["Duração"], errors="coerce").sum()
    ponderado = pd.to_numeric(df["Progresso Ponderado"], errors="coerce").sum()
    return ponderado / duracao if duracao else 0


try:
    entregas, marcos, origem = carregar_dados()
except Exception as erro:
    st.error("Não foi possível carregar a planilha. Confirme se o compartilhamento está como ‘qualquer pessoa com o link’. ")
    st.exception(erro)
    st.stop()

hoje = pd.Timestamp.now().normalize()

with st.sidebar:
    st.markdown("### Filtros")
    fases = st.multiselect("Fase", sorted(entregas["Fase"].unique()), default=[])
    frentes = st.multiselect("Frente", sorted(entregas["Frente"].unique()), default=[])
    tipos = st.multiselect("Tipo", sorted(entregas["Tipo"].unique()), default=[])
    status = st.multiselect("Status", sorted(entregas["Status"].unique()), default=[])
    datas_validas = entregas["Data"].dropna()
    if not datas_validas.empty:
        periodo = st.date_input("Período", value=(datas_validas.min().date(), datas_validas.max().date()), min_value=datas_validas.min().date(), max_value=datas_validas.max().date())
    else:
        periodo = ()
    st.divider()
    if st.button("Atualizar dados", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.caption(f"Fonte: {origem} · cache de 10 min")

df = entregas.copy()
if fases:
    df = df[df["Fase"].isin(fases)]
if frentes:
    df = df[df["Frente"].isin(frentes)]
if tipos:
    df = df[df["Tipo"].isin(tipos)]
if status:
    df = df[df["Status"].isin(status)]
if isinstance(periodo, (tuple, list)) and len(periodo) == 2:
    inicio, fim = pd.Timestamp(periodo[0]), pd.Timestamp(periodo[1])
    df = df[df["Data"].isna() | df["Data"].between(inicio, fim)]

st.markdown(
    """<div class="hero"><h1>Painel Executivo do Projeto SESAP</h1>
    <p>Entregas, resultados e evolução dos marcos · UFRN + Secretaria de Estado da Saúde Pública do RN</p></div>""",
    unsafe_allow_html=True,
)

abas = st.tabs(["Visão executiva", "Entregas e resultados", "Evolução dos marcos", "Governança dos dados"])

with abas[0]:
    total = len(df)
    realizados = int((df["Status"] == "Realizado").sum())
    pendentes = int((df["Status"] == "Pendente").sum())
    taxa = realizados / total if total else 0
    progresso = progresso_ponderado(marcos)
    ultima = df["Data"].max()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Entregas realizadas", f"{realizados}")
    c2.metric("Taxa de conclusão", f"{taxa:.1%}")
    c3.metric("Pendências", f"{pendentes}")
    c4.metric("Progresso dos marcos", f"{progresso:.1f}%")
    c5.metric("Capacidade semanal", "160 h")

    st.markdown('<div class="section-title">Jornada do projeto</div>', unsafe_allow_html=True)
    periodos = [
        ("Fase 1 · Estratégica", "out/2024 — set/2025", pd.Timestamp("2024-10-01"), pd.Timestamp("2025-09-30")),
        ("Fase 2 · Tática", "out/2025 — set/2026", pd.Timestamp("2025-10-01"), pd.Timestamp("2026-09-30")),
        ("Fase 3 · Operacional", "out/2026 — out/2027", pd.Timestamp("2026-10-01"), pd.Timestamp("2027-10-31")),
    ]
    cols = st.columns(3)
    for col, (nome, periodo_txt, ini, fim) in zip(cols, periodos):
        if hoje > fim:
            classe, rotulo = "done", "Concluída"
        elif hoje >= ini:
            classe, rotulo = "active", "Em execução"
        else:
            classe, rotulo = "future", "Planejada"
        extra = " phase-active" if classe == "active" else ""
        col.markdown(f'<div class="phase-card{extra}"><div class="phase-name">{nome}</div><div class="phase-period">{periodo_txt}</div><span class="phase-status {classe}">{rotulo}</span></div>', unsafe_allow_html=True)

    esquerda, direita = st.columns([1.45, 1])
    mensal = df.dropna(subset=["Mês"]).groupby(["Mês", "Status"]).size().reset_index(name="Entregas")
    fig = px.bar(mensal, x="Mês", y="Entregas", color="Status", barmode="stack", title="Evolução mensal das entregas", color_discrete_map={"Realizado": CORES["verde_agua"], "Pendente": CORES["amarelo"]})
    esquerda.plotly_chart(grafico_layout(fig), use_container_width=True)

    por_fase = df.groupby(["Fase", "Status"]).size().reset_index(name="Entregas")
    fig = px.bar(por_fase, y="Fase", x="Entregas", color="Status", orientation="h", title="Situação por fase", color_discrete_map={"Realizado": CORES["verde"], "Pendente": CORES["amarelo"]})
    direita.plotly_chart(grafico_layout(fig), use_container_width=True)

    st.caption(f"Último registro datado: {ultima.strftime('%d/%m/%Y') if pd.notna(ultima) else 'não disponível'} · 11 integrantes · capacidade nominal de 160 h/semana")

with abas[1]:
    realizados_df = df[df["Status"] == "Realizado"]
    contagens = realizados_df["Tipo"].value_counts()
    cols = st.columns(4)
    for col, tipo in zip(cols, ["Produto", "Reunião", "Visitas Técnicas", "Capacitação"]):
        col.metric(tipo, int(contagens.get(tipo, 0)))

    c1, c2 = st.columns(2)
    frente = df.groupby(["Frente", "Status"]).size().reset_index(name="Entregas")
    fig = px.bar(frente, y="Frente", x="Entregas", color="Status", orientation="h", title="Entregas por frente", color_discrete_map={"Realizado": CORES["verde_agua"], "Pendente": CORES["amarelo"]})
    c1.plotly_chart(grafico_layout(fig, 470), use_container_width=True)
    subtipo = realizados_df[realizados_df["Subtipo"] != "Não informado"]["Subtipo"].value_counts().head(12).sort_values().reset_index()
    fig = px.bar(subtipo, y="Subtipo", x="count", orientation="h", title="Principais produtos e atividades", color_discrete_sequence=[CORES["azul"]])
    fig.update_xaxes(title="Entregas"); fig.update_yaxes(title="")
    c2.plotly_chart(grafico_layout(fig, 470), use_container_width=True)

    st.markdown('<div class="section-title">Detalhamento das entregas</div>', unsafe_allow_html=True)
    tabela = df[["Data", "Fase", "Tipo", "Subtipo", "Descrição", "Frente", "Responsável", "Local", "Status", "Lista Mestra"]].sort_values("Data", ascending=False)
    st.dataframe(tabela, use_container_width=True, hide_index=True, column_config={"Data": st.column_config.DateColumn("Data", format="DD/MM/YYYY")}, height=440)
    st.download_button("Baixar dados filtrados (CSV)", tabela.to_csv(index=False).encode("utf-8-sig"), "entregas_sesap.csv", "text/csv")

with abas[2]:
    fase_marco = st.selectbox("Fase dos marcos", marcos["Fase"].unique())
    mf = marcos[marcos["Fase"] == fase_marco].copy()
    prog = progresso_ponderado(mf)
    concluido = int((mf["Situação"] == "Concluído").sum())
    andamento = int((mf["Situação"] == "Em andamento").sum())
    nao_iniciado = int((mf["Situação"] == "Não iniciado").sum())
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Progresso ponderado", f"{prog:.1f}%")
    c2.metric("Concluídos", concluido)
    c3.metric("Em andamento", andamento)
    c4.metric("Não iniciados", nao_iniciado)

    fig = px.timeline(mf, x_start="Início", x_end="Fim", y="Marco", color="Situação", title="Cronograma dos marcos", color_discrete_map={"Concluído": CORES["verde"], "Em andamento": CORES["amarelo"], "Não iniciado": "#B8C4C7"})
    fig.update_yaxes(autorange="reversed")
    st.plotly_chart(grafico_layout(fig, 430), use_container_width=True)

    fig = go.Figure()
    fig.add_trace(go.Bar(y=mf["Marco"], x=[100] * len(mf), orientation="h", marker_color="#E8EEF0", hoverinfo="skip", name="Restante"))
    fig.add_trace(go.Bar(y=mf["Marco"], x=mf["Progresso"], orientation="h", marker_color=CORES["verde_agua"], text=mf["Progresso"].map(lambda x: f"{x:.0f}%"), textposition="inside", name="Progresso"))
    fig.update_layout(barmode="overlay", title="Progresso de cada marco", xaxis_range=[0, 100], xaxis_title="Percentual")
    st.plotly_chart(grafico_layout(fig, 430), use_container_width=True)

with abas[3]:
    sem_data = int(entregas["Data"].isna().sum())
    sem_resp = int((entregas["Responsável"] == "Não informado").sum())
    sem_horas = int(entregas["Horas"].isna().sum())
    lista_pendente = int(entregas["Lista Mestra"].isin(["Pendente", "A realizar"]).sum())
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Registros sem data", sem_data)
    c2.metric("Sem responsável", sem_resp)
    c3.metric("Sem horas", sem_horas)
    c4.metric("Lista mestra pendente", lista_pendente)
    st.markdown('<div class="note">As 64 horas registradas não representam o esforço total do projeto: somente 8 dos 147 registros possuem essa informação. A capacidade de 160 h/semana é nominal e não deve ser comparada diretamente com esse campo.</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Registros que precisam de saneamento</div>', unsafe_allow_html=True)
    problemas = entregas[entregas["Data"].isna() | (entregas["Responsável"] == "Não informado") | entregas["Horas"].isna()]
    st.dataframe(problemas[["ID", "Data", "Descrição", "Frente", "Responsável", "Horas", "Status"]], use_container_width=True, hide_index=True, height=460)

