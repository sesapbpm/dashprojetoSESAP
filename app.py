from io import BytesIO
from pathlib import Path
import os
import base64

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from PIL import Image


st.set_page_config(
    page_title="Projeto SESAP | Painel Executivo",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
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
    .hero { padding: 22px 26px; border-radius: 16px; color: white; margin-bottom: 18px;
            background: linear-gradient(110deg, #004E61 0%, #007D86 60%, #00A6A6 100%); }
    .hero-grid { display:grid; grid-template-columns:190px 1fr; gap:28px; align-items:center; }
    .hero-logo { width:190px; height:150px; object-fit:contain; background:transparent; padding:0; }
    .hero h1 { margin: 0; font-size: 1.62rem; line-height:1.2; font-weight: 700; }
    .hero p { margin: 9px 0 0; color: #E6F7F7; line-height:1.45; }
    .section-title { font-size: 1.15rem; font-weight: 700; color: #17313A; margin: 12px 0 8px; }
    .phase-card { background: white; border: 1px solid #DDE6E8; border-radius: 12px; padding: 14px 16px; min-height: 112px; }
    .phase-active { border: 2px solid #00A6A6; box-shadow: 0 3px 12px rgba(0,166,166,.12); }
    .phase-name { font-weight: 700; color: #17313A; }
    .phase-period { font-size: .86rem; color: #60747B; margin-top: 4px; }
    .phase-status { display:inline-block; margin-top:12px; padding:4px 9px; border-radius:99px; font-size:.78rem; font-weight:600; }
    .done { background:#E5F5E9; color:#238636; } .active { background:#DDF5F3; color:#006D72; }
    .future { background:#EEF1F2; color:#60747B; }
    .note { background:#FFF9E8; border-left:4px solid #F2B134; padding:12px 14px; border-radius:6px; color:#5C4A12; }
    .capacity { background:linear-gradient(90deg,#EAF6F6,#FFFFFF); border:1px solid #CDE7E6; border-radius:14px; padding:16px 20px; margin:12px 0 18px; }
    .result-card { background:#FFFFFF; border:1px solid #DDE6E8; border-radius:14px; padding:18px; min-height:190px; }
    .result-card h3 { font-size:1rem; margin:0 0 10px; color:#005A70; }
    .result-card p, .result-card li { font-size:.9rem; color:#52646B; line-height:1.45; }
    .impact { background:#073670; border-radius:14px; padding:20px; color:white; }
    .impact h3 { margin-top:0; color:white; } .impact li { margin:7px 0; color:#EDF7FF; }
    @media(max-width:700px){ .hero-grid{grid-template-columns:1fr}.hero-logo{width:150px;height:115px}.hero h1{font-size:1.25rem} }
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
    if (ROOT / "LogoProjeto.png").exists():
        logo_sidebar = Image.open(ROOT / "LogoProjeto.png")
        if logo_sidebar.mode == "RGBA" and logo_sidebar.getbbox():
            logo_sidebar = logo_sidebar.crop(logo_sidebar.getbbox())
        st.image(logo_sidebar, use_container_width=True)
    st.caption("DIMP · SESAP/RN · UFRN")
    st.divider()
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
    st.markdown("#### Área interna")
    try:
        chave_configurada = os.getenv("INTERNAL_ACCESS_KEY") or st.secrets.get("INTERNAL_ACCESS_KEY", "")
    except Exception:
        chave_configurada = os.getenv("INTERNAL_ACCESS_KEY", "")
    chave_informada = st.text_input("Chave da equipe", type="password", help="Libera indicadores de qualidade e saneamento da base.")
    acesso_interno = bool(chave_configurada and chave_informada == chave_configurada)
    if chave_informada and not acesso_interno:
        st.error("Chave interna inválida.")
    if acesso_interno:
        st.success("Modo interno habilitado")
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

logo_path = ROOT / "LogoProjeto.png"
if logo_path.exists():
    logo_imagem = Image.open(logo_path)
    if logo_imagem.mode == "RGBA" and logo_imagem.getbbox():
        logo_imagem = logo_imagem.crop(logo_imagem.getbbox())
    logo_buffer = BytesIO()
    logo_imagem.save(logo_buffer, format="PNG")
    logo_b64 = base64.b64encode(logo_buffer.getvalue()).decode()
else:
    logo_b64 = ""
logo_html = f'<img class="hero-logo" src="data:image/png;base64,{logo_b64}" alt="Logo DIMP-SESAP">' if logo_b64 else ""
st.markdown(f"""
    <div class="hero"><div class="hero-grid">{logo_html}<div>
    <h1>Desenvolvimento e Implantação de Ferramentas BPM</h1>
    <p>Modernização dos processos de gestão nos níveis estratégico, tático e operacional da Secretaria de Estado da Saúde Pública do Rio Grande do Norte — SESAP/RN.</p>
    <p><b>Parceria:</b> SESAP/RN e Departamento de Engenharia de Produção da UFRN</p>
    </div></div></div>""", unsafe_allow_html=True)

nomes_abas = ["Visão executiva", "Entregas e resultados", "Evolução dos marcos", "Resultados das intervenções"]
if acesso_interno:
    nomes_abas.append("🔒 Governança dos dados")
abas = st.tabs(nomes_abas)

with abas[0]:
    total = len(df)
    realizados = int((df["Status"] == "Realizado").sum())
    pendentes = int((df["Status"] == "Pendente").sum())
    taxa = realizados / total if total else 0
    progresso = progresso_ponderado(marcos)
    ultima = df["Data"].max()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Entregas realizadas", f"{realizados}")
    c2.metric("Taxa de conclusão", f"{taxa:.1%}")
    c3.metric("Pendências", f"{pendentes}")
    c4.metric("Progresso dos marcos", f"{progresso:.1f}%")

    inicio_projeto = pd.Timestamp("2024-10-01")
    fim_projeto = pd.Timestamp("2027-10-31")
    data_referencia = min(max(hoje, inicio_projeto), fim_projeto)
    dias_uteis_decorridos = len(pd.bdate_range(inicio_projeto, data_referencia))
    horas_acumuladas = round(dias_uteis_decorridos / 5 * 160)
    dias_pessoa_acumulados = round(horas_acumuladas / 8)
    total_dias_uteis = len(pd.bdate_range(inicio_projeto, fim_projeto))
    total_horas_projeto = round(total_dias_uteis / 5 * 160)
    total_dias_pessoa = round(total_horas_projeto / 8)
    st.markdown('<div class="section-title">Capacidade de dedicação da equipe</div>', unsafe_allow_html=True)
    h1, h2, h3, h4 = st.columns(4)
    h1.metric("Capacidade semanal", "160 h", "20 dias-pessoa")
    h2.metric("Equipe dedicada", "11 pessoas", "5 professores · 5 bolsistas · 1 gerente")
    h3.metric("Capacidade acumulada", f"{horas_acumuladas:,.0f} h".replace(",", "."), f"{dias_pessoa_acumulados:,.0f} dias-pessoa".replace(",", "."))
    h4.metric("Capacidade total planejada", f"{total_horas_projeto:,.0f} h".replace(",", "."), f"{total_dias_pessoa:,.0f} dias-pessoa".replace(",", "."))
    st.caption("Capacidade nominal: professores 8 h/semana; bolsistas e gerente 20 h/semana. Dias-pessoa calculados a 8 horas por dia, sem desconto de feriados e afastamentos.")

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
    capacitacoes = entregas[(entregas["Tipo"] == "Capacitação") & (entregas["Status"] == "Realizado")]
    horas_capacitacao = capacitacoes["Horas"].sum()
    treinamentos_com_hora = int(capacitacoes["Horas"].notna().sum())
    st.markdown('<div class="section-title">Principais resultados e ações desenvolvidas</div>', unsafe_allow_html=True)
    st.caption("Resultados consolidados das intervenções realizadas nas fases Estratégica e Tática.")

    r1, r2 = st.columns(2)
    r1.markdown('''<div class="result-card"><h3>1. Estruturação e governança</h3><ul>
    <li><b>Estruturação da DPCON:</b> proposta de criação da Diretoria de Processos e Contratos, com competências, macroprocessos e modelo de governança.</li>
    <li><b>Layout da DPCON:</b> proposta de espaço físico otimizado, integrando núcleos e melhorando a comunicação interna.</li>
    <li><b>Pontos de remuneração:</b> critérios e modelo de distribuição para fiscais de contrato, com foco em desempenho e equidade.</li>
    </ul></div>''', unsafe_allow_html=True)
    r2.markdown('''<div class="result-card"><h3>2. Contratos e manutenção</h3><ul>
    <li><b>Minuta para contratos de manutenção:</b> modelo padronizado baseado em Planejamento e Controle de Manutenção — PCM.</li>
    <li><b>Contratos.gov:</b> adoção da plataforma para publicação, acompanhamento e gestão centralizada dos contratos.</li>
    </ul></div>''', unsafe_allow_html=True)

    r3, r4 = st.columns(2)
    r3.markdown('''<div class="result-card"><h3>3. Central de Compras</h3><p>Diagnóstico detalhado da situação atual, identificação de gargalos, riscos e oportunidades e proposição de estrutura e processos para consolidação da Central de Compras da SESAP.</p></div>''', unsafe_allow_html=True)
    r4.markdown('''<div class="result-card"><h3>4. Custos e contabilidade</h3><p>Mapeamento e análise dos custos dos serviços de saúde no sistema APURASUS, ampliando a precisão do custeio e o suporte à tomada de decisão baseada em dados.</p></div>''', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Automação e ganho de eficiência</div>', unsafe_allow_html=True)
    a1, a2 = st.columns(2)
    a1.metric("Robô de aquisições", "> 90%", "redução no tempo de extração e consolidação")
    a1.caption("Automação via RPA da extração de dados do SIPAC, antes realizada manualmente.")
    a2.metric("Planilha de terceirizados", "> 99%", "redução no tempo de análise")
    a2.caption("Conferência automatizada das folhas de ponto: de aproximadamente 57h40 para poucos minutos.")

    p1, p2 = st.columns(2)
    p1.markdown('''<div class="result-card"><h3>6. Produção científica</h3><ul>
    <li>Submissão de artigos científicos e apresentações em eventos.</li><li>Desenvolvimento de TCCs e iniciação científica.</li>
    <li>Relatórios técnicos, ferramentas e materiais metodológicos.</li></ul></div>''', unsafe_allow_html=True)
    p2.markdown(f'''<div class="result-card"><h3>7. Treinamentos realizados</h3>
    <p style="font-size:1.8rem;font-weight:700;color:#B27800;margin:4px 0">{horas_capacitacao:.0f} h</p>
    <p>de capacitações executadas em <b>{treinamentos_com_hora}</b> registros com carga horária informada.</p>
    <p>Formato presencial e on-line, alcançando fiscais e gestores da SESAP. Temas: BPM, SIPAC, PCM, gestão de contratos, planilhas, indicadores e ferramentas de gestão.</p></div>''', unsafe_allow_html=True)

    st.markdown('''<div class="impact"><h3>Impactos gerados</h3><ul>
    <li>Mais eficiência e agilidade nos processos.</li><li>Maior controle, transparência e rastreabilidade das informações.</li>
    <li>Redução de retrabalho e erros manuais.</li><li>Melhor gestão dos contratos e dos recursos públicos.</li>
    <li>Base mais sólida para decisões estratégicas orientadas por dados.</li></ul></div>''', unsafe_allow_html=True)

if acesso_interno:
  with abas[4]:
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
