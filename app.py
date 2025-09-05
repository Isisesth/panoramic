import io
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, date, timedelta
from docx import Document

# =========================
# CONFIGURAÇÕES GERAIS
# =========================
st.set_page_config(
    page_title="USA4ALL • Panorama",
    layout="wide",
    page_icon="🗂️",
)

# Paleta (verde escuro + verde-lima)
PRIMARY = "#0B3D2E"   # verde escuro
ACCENT  = "#9BE84F"   # verde-lima
BG_SOFT = "#0F4737"

# Logo (URL direto do thumb do YouTube repassado)
LOGO_URL = "https://i.ytimg.com/vi/aWcON7jyz0I/hq720.jpg"

# CSS para tema
st.markdown(
    f"""
    <style>
    .stApp {{
        background: linear-gradient(180deg, {BG_SOFT} 0%, #07281F 100%);
        color: #ffffff;
    }}
    header, .st-emotion-cache-18ni7ap, .st-emotion-cache-12fmjuu {{
        background-color: transparent !important;
    }}
    section[data-testid="stSidebar"] > div {{
        background: #0A3327;
        color: #fff;
        border-right: 1px solid rgba(255,255,255,0.1);
    }}
    .stButton>button, .stDownloadButton>button {{
        background: {PRIMARY} !important;
        color: #fff !important;
        border: 1px solid {ACCENT} !important;
        border-radius: 8px !important;
    }}
    .stMetric-value, .stMetric-label {{
        color: #ffffff !important;
    }}
    .stProgress > div > div > div > div {{
        background-color: {ACCENT} !important;
    }}
    .stSelectbox div[data-baseweb="select"] > div {{
        background: #123F30;
        color: #fff;
    }}
    .stDataFrame, .stTable {{
        background: rgba(255,255,255,0.03);
    }}
    </style>
    """,
    unsafe_allow_html=True
)

# =========================
# FUNÇÕES AUXILIARES
# =========================
def parse_date(val, fallback=None):
    """Converte para date com dayfirst=True (dd/mm/aaaa)."""
    if pd.isna(val) or val is None or str(val).strip() == "":
        return fallback
    try:
        d = pd.to_datetime(val, errors="coerce", dayfirst=True)
        if pd.isna(d):
            return fallback
        return d.date()
    except Exception:
        return fallback

def fmt_date(d):
    return d.strftime("%d/%m/%Y") if isinstance(d, (datetime, date)) else "—"

def safe_progress_value(percentual):
    return max(0.0, min((percentual or 0.0)/100.0, 1.0))

def df_to_csv_bytes(df: pd.DataFrame, include_index: bool = True) -> bytes:
    return df.to_csv(index=include_index).encode("utf-8-sig")

# Prazos SOL por Practice Area
SOL_PRAZO = {
    "FOIA": 30, "I-130": 30, "COS": 30, "B2-EXT": 30, "NPT": 60, "NVC": 60, "K1": 30,
    "WAIVER": 90, "EB2-NIW": 120, "EB1": 90, "E2": 90, "O1": 90, "EB4": 90,
    "AOS": 60, "ROC": 60, "NATZ": 30, "DA": 90, "PIP": 30, "I-90": 30, "COURT": 60,
    "DOM": 30, "SIJS": 30, "VAWA": 90, "T-VISA": 120, "U-VISA": 90, "I-918B": 30,
    "ASYLUM": 120, "PERM": 30, "EB3": 30
}

# =========================
# SIDEBAR • LOGO + CONTROLES
# =========================
with st.sidebar:
    st.image(LOGO_URL, caption="USA4ALL", use_column_width=True)
    st.markdown("## ⚙️ Modo de uso")
    mode = st.radio("Preenchimento", ["A partir de arquivo", "Manual"], horizontal=False)
    st.markdown("---")
    st.markdown("**Observação:** para ver **duração por estágio**, envie também o arquivo **Histórico de Estágios**.")

st.title("🗂️ Panorama de Casos — USA4ALL")

# =========================
# ESTADO DE SESSÃO
# =========================
if "courses" not in st.session_state:
    st.session_state.courses = [{"curso": "", "universidade": "", "conclusao": datetime.now().date()}]
if "stages_manual" not in st.session_state:
    st.session_state.stages_manual = []
if "df_cases" not in st.session_state:
    st.session_state.df_cases = None
if "df_stages" not in st.session_state:
    st.session_state.df_stages = None

# =========================
# MAPEAMENTO DE COLUNAS (CASES + STAGES)
# =========================
CASES_FIELDS = {
    "Case": ["Case","Caso","Cliente","Assunto"],
    "Case Number": ["Case Number","Número do Caso","CaseNo","Case_ID","ID"],
    "Practice Area": ["Practice Area","Área","Area","Tipo de Visto","Visto"],
    "Case Stage": ["Case Stage","Stage","Status","Fase","Etapa"],
    "Open Date": ["Open Date","Data de Abertura","Start Date","Início"],
    "Closed Date": ["Closed Date","Data de Fechamento","End Date","Fechado em"],
    "Statute of Limitations Date": ["Statute of Limitations Date","SOL","SOL Date","Prazo SOL","Limitation Date"],
}
STAGES_FIELDS = {
    "Case Number": ["Case Number","Número do Caso","CaseNo","ID"],
    "Case Stage": ["Case Stage","Stage","Status","Fase","Etapa"],
    "Start Date": ["Start Date","Início","Data Inicial","Start"],
    "End Date": ["End Date","Fim","Data Final","End"],
}

def suggest_mapping(df_cols, synonyms):
    norm = {c: c.strip().lower() for c in df_cols if isinstance(c, str)}
    # Exato
    for syn in synonyms:
        s = syn.strip().lower()
        for c, cl in norm.items():
            if cl == s:
                return c
    # Contém
    for syn in synonyms:
        s = syn.strip().lower()
        for c, cl in norm.items():
            if s in cl:
                return c
    return None

def mapping_ui(df, expected_dict, title):
    st.markdown(f"#### 🔎 Mapeamento de colunas — {title}")
    cols = list(df.columns)
    options = ["(não usar)"] + cols
    mapping = {}
    left, right = st.columns(2)
    keys = list(expected_dict.keys())
    half = len(keys)//2
    with left:
        for k in keys[:half]:
            sug = suggest_mapping(cols, expected_dict[k]) or "(não usar)"
            mapping[k] = st.selectbox(k, options, index=options.index(sug), key=f"map_{title}_{k}")
    with right:
        for k in keys[half:]:
            sug = suggest_mapping(cols, expected_dict[k]) or "(não usar)"
            mapping[k] = st.selectbox(k, options, index=options.index(sug), key=f"map_{title}_{k}")

    rename = {v: k for k, v in mapping.items() if v != "(não usar)"}
    df2 = df.rename(columns=rename).copy()
    # Trim e datas
    for c in df2.columns:
        if isinstance(c, str):
            df2[c] = df2[c].apply(lambda x: x.strip() if isinstance(x, str) else x)
    # Datas conhecidas
    for dc in ["Open Date","Closed Date","Statute of Limitations Date","Start Date","End Date"]:
        if dc in df2.columns:
            df2[dc] = pd.to_datetime(df2[dc], errors="coerce", dayfirst=True).dt.date
    st.success("✔️ Mapeamento aplicado.")
    return df2

# =========================
# UPLOAD / ENTRADA DE DADOS
# =========================
df_cases = None
df_stages = None
selected_case = None

if mode == "A partir de arquivo":
    st.subheader("📂 Upload de dados")
    up1 = st.file_uploader("Arquivo de **Casos** (CSV/XLS/XLSX)", type=["csv","xls","xlsx"], key="up_cases")
    up2 = st.file_uploader("Arquivo **Histórico de Estágios** (opcional) — colunas: Case Number, Case Stage, Start Date, End Date",
                           type=["csv","xls","xlsx"], key="up_stages")

    if up1:
        try:
            if up1.name.lower().endswith((".xls",".xlsx")):
                raw_cases = pd.read_excel(up1)
            else:
                raw_cases = pd.read_csv(up1)
            with st.expander("Ajustar colunas (Casos)"):
                df_cases = mapping_ui(raw_cases, CASES_FIELDS, "Casos")
            st.session_state.df_cases = df_cases
            st.success("✅ Casos carregados.")
            st.dataframe(df_cases.head(50), use_container_width=True)
        except Exception as e:
            st.error(f"Erro ao ler Casos: {e}")

    if up2:
        try:
            if up2.name.lower().endswith((".xls",".xlsx")):
                raw_stages = pd.read_excel(up2)
            else:
                raw_stages = pd.read_csv(up2)
            with st.expander("Ajustar colunas (Histórico de Estágios)"):
                df_stages = mapping_ui(raw_stages, STAGES_FIELDS, "Estágios")
            st.session_state.df_stages = df_stages
            st.success("✅ Histórico de Estágios carregado.")
            st.dataframe(df_stages.head(50), use_container_width=True)
        except Exception as e:
            st.error(f"Erro ao ler Estágios: {e}")

    # Persistência na sessão
    if st.session_state.df_cases is not None:
        df_cases = st.session_state.df_cases
    if st.session_state.df_stages is not None:
        df_stages = st.session_state.df_stages

    # Seletor de Case
    if df_cases is not None and "Case Number" in df_cases.columns:
        st.subheader("🔎 Selecione um cliente (Case)")
        cases_list = df_cases["Case Number"].dropna().astype(str).unique().tolist()
        if cases_list:
            selected_case = st.selectbox("Case Number", cases_list)
        else:
            st.warning("Nenhum 'Case Number' encontrado.")

    # ======= PAINEL DO CASE SELECIONADO =======
    if selected_case and df_cases is not None:
        cdata = df_cases[df_cases["Case Number"].astype(str) == str(selected_case)].iloc[0].to_dict()
        nome          = cdata.get("Case","")
        practice_area = cdata.get("Practice Area","")
        stage_current = cdata.get("Case Stage","")
        open_date     = parse_date(cdata.get("Open Date"))
        closed_date   = parse_date(cdata.get("Closed Date"))
        sol_date      = parse_date(cdata.get("Statute of Limitations Date"))

        st.markdown(f"### 📌 {practice_area or '—'} • **{nome or '—'}**  \n**Case Number:** {selected_case}")

        # Progresso vs SOL
        hoje = datetime.now().date()
        if open_date and sol_date:
            tot = (sol_date - open_date).days
            dec = (hoje - open_date).days
            rest = (sol_date - hoje).days
            perc = round((dec/tot)*100,2) if tot>0 else (100.0 if hoje>=sol_date else 0.0)
        else:
            tot=dec=rest=0
            perc = 0.0
        col_a, col_b, col_c = st.columns(3)
        col_a.metric("Dias decorridos", f"{max(dec,0)}")
        col_b.metric("Dias até SOL", f"{max(rest,0)}")
        col_c.metric("Progresso", f"{perc:.2f}%")
        st.progress(safe_progress_value(perc))

        if open_date and sol_date:
            if rest < 0:
                st.error("⚠️ SOL ultrapassado.")
            elif rest <= 5:
                st.warning("⏱️ Menos de 5 dias para o SOL.")
            else:
                st.success("✔️ Dentro do prazo do SOL.")
        else:
            st.info("Datas insuficientes para avaliar SOL.")

        # --- GRÁFICO: duração por Case Stage (APENAS QUANDO UM CLIENTE É SELECIONADO)
        st.subheader("⏱️ Duração por Case Stage (Case selecionado)")
        # Monta histórico deste case (de df_stages)
        if df_stages is not None and all(col in df_stages.columns for col in ["Case Number","Case Stage","Start Date","End Date"]):
            h = df_stages[df_stages["Case Number"].astype(str)==str(selected_case)].copy()
            if not h.empty:
                # Calcula dias por linha
                def _dur(r):
                    sd = parse_date(r.get("Start Date"))
                    ed = parse_date(r.get("End Date"), fallback=hoje)  # se fim vazio, usa hoje
                    if sd and ed and ed >= sd:
                        return (ed - sd).days
                    return 0
                h["Dias"] = h.apply(_dur, axis=1)
                # Ordena por início
                h["Start Date"] = pd.to_datetime(h["Start Date"], errors="coerce", dayfirst=True).dt.date
                h = h.sort_values("Start Date")

                # Plot (barras horizontais)
                fig_h = max(3, 0.5 * len(h))
                fig, ax = plt.subplots(figsize=(10, fig_h))
                y = range(len(h))
                ax.barh(y, h["Dias"].tolist(), color=ACCENT)
                ax.set_yticks(list(y))
                labels = [
                    f"{row['Case Stage']} ({fmt_date(row['Start Date'])} → {fmt_date(parse_date(row['End Date'], fallback=hoje))})"
                    for _, row in h.iterrows()
                ]
                ax.set_yticklabels(labels)
                ax.invert_yaxis()
                ax.set_xlabel("Dias")
                ax.set_title("Tempo em cada Stage")
                fig.patch.set_facecolor(BG_SOFT)
                ax.set_facecolor("#0B2C21")
                st.pyplot(fig, clear_figure=True)
            else:
                st.info("Não há histórico de estágios para este case no arquivo enviado.")
        else:
            st.info("Envie o arquivo **Histórico de Estágios** para ver a duração por Stage.")

# =========================
# MODO MANUAL (continua disponível)
# =========================
if mode == "Manual":
    st.subheader("📝 Cadastro rápido (Manual)")
    # Área e datas
    tipo = st.selectbox("Practice Area", list(SOL_PRAZO.keys()))
    data_inicio = st.date_input("Data de início", value=datetime.now().date(), format="DD/MM/YYYY")
    sol_dias = SOL_PRAZO[tipo]
    prazo_final = data_inicio + timedelta(days=sol_dias)
    hoje = datetime.now().date()
    tot = (prazo_final - data_inicio).days
    dec = (hoje - data_inicio).days
    rest = (prazo_final - hoje).days
    perc = round((dec/tot)*100,2) if tot>0 else (100.0 if hoje>=prazo_final else 0.0)

    c1,c2,c3 = st.columns(3)
    c1.metric("Dias decorridos", f"{max(dec,0)}")
    c2.metric("Dias até SOL", f"{max(rest,0)}")
    c3.metric("Progresso", f"{perc:.2f}%")
    st.progress(safe_progress_value(perc))

    # Estágios manuais
    st.markdown("#### Estágios")
    if st.button("Adicionar estágio (manual)"):
        base = st.session_state.stages_manual[-1]["end"] if st.session_state.stages_manual else data_inicio
        st.session_state.stages_manual.append({"stage":"(defina)", "start":base, "end":base})
    for i, s in enumerate(st.session_state.stages_manual):
        st.session_state.stages_manual[i]["stage"] = st.text_input(f"Stage {i+1}", value=s["stage"], key=f"m_stage_{i}")
        st.session_state.stages_manual[i]["start"] = st.date_input("Início", value=s["start"], key=f"m_start_{i},", format="DD/MM/YYYY")
        st.session_state.stages_manual[i]["end"] = st.date_input("Fim", value=s["end"], key=f"m_end_{i}", format="DD/MM/YYYY")
        d = (st.session_state.stages_manual[i]["end"] - st.session_state.stages_manual[i]["start"]).days
        st.caption(f"⏳ {max(d,0)} dias")

# =========================
# OVERVIEW DO DEPARTAMENTO (por tipo de visto / Practice Area)
# =========================
st.subheader("🏢 Overview do Departamento — por Practice Area")
df_cases_used = st.session_state.df_cases

if df_cases_used is not None and not df_cases_used.empty and "Practice Area" in df_cases_used.columns:
    dfc = df_cases_used.copy()
    dfc["Practice Area"] = dfc["Practice Area"].astype(str).str.strip()
    # Ativos = exclude Approved/Denied no Stage
    def is_active(stage):
        s = str(stage).upper()
        return not(("APPROVED" in s) or ("DENIED" in s) or ("CLOSED" in s))
    dfc["Ativo"] = dfc["Case Stage"].apply(is_active)

    resumo = dfc.groupby("Practice Area").agg(
        total=("Case","count"),
        ativos=("Ativo","sum")
    ).sort_values("ativos", ascending=False).reset_index()

    st.dataframe(resumo, use_container_width=True)

    # Gráfico barras (ativos por área)
    fig, ax = plt.subplots(figsize=(10, max(3, 0.5*len(resumo))))
    ax.barh(resumo["Practice Area"], resumo["ativos"], color=ACCENT)
    ax.invert_yaxis()
    ax.set_xlabel("Casos Ativos")
    ax.set_title("Casos Ativos por Practice Area")
    fig.patch.set_facecolor(BG_SOFT)
    ax.set_facecolor("#0B2C21")
    st.pyplot(fig, clear_figure=True)

else:
    st.caption("Carregue o arquivo de **Casos** para ver o overview por área.")

# =========================
# DIAS MÉDIOS POR CASE STAGE (GERAL, independente da área)
# =========================
st.subheader("📊 Dias médios por Case Stage (geral)")
df_stages_used = st.session_state.df_stages

if df_stages_used is not None and not df_stages_used.empty and all(c in df_stages_used.columns for c in ["Case Stage","Start Date","End Date"]):
    dfg = df_stages_used.copy()

    # Duração por linha
    def dur(r):
        sd = parse_date(r.get("Start Date"))
        ed = parse_date(r.get("End Date"), fallback=datetime.now().date())
        if sd and ed and ed >= sd:
            return (ed - sd).days
        return 0
    dfg["Dias"] = dfg.apply(dur, axis=1)

    media_por_stage = dfg.groupby("Case Stage")["Dias"].mean().round(1).sort_values(ascending=False)
    media_df = media_por_stage.reset_index().rename(columns={"Dias":"Dias médios"})
    st.dataframe(media_df, use_container_width=True)

    # Gráfico
    fig, ax = plt.subplots(figsize=(10, max(3, 0.45*len(media_df))))
    ax.barh(media_df["Case Stage"], media_df["Dias médios"], color=ACCENT)
    ax.invert_yaxis()
    ax.set_xlabel("Dias médios")
    ax.set_title("Dias médios por Stage (geral)")
    fig.patch.set_facecolor(BG_SOFT); ax.set_facecolor("#0B2C21")
    st.pyplot(fig, clear_figure=True)
else:
    st.caption("Envie o **Histórico de Estágios** para calcular as médias por stage.")

# =========================
# ESTIMATIVA: TEMPO MÉDIO DE CONCLUSÃO POR ÁREA
# (exclui períodos 'USCIS Pending Decision')
# e MÉDIA DE DIAS ULTRAPASSADOS DO SOL
# =========================
st.subheader("⏳ Estimativa — Tempo médio de conclusão por Área (excluindo 'USCIS Pending Decision') + SOL")

if (df_cases_used is not None and "Open Date" in df_cases_used.columns) and (df_stages_used is not None):
    cases = df_cases_used.copy()
    stages = df_stages_used.copy()
    # Normalizações
    for c in ["Open Date","Closed Date","Statute of Limitations Date"]:
        if c in cases.columns:
            cases[c] = cases[c].apply(parse_date)
    for c in ["Start Date","End Date"]:
        if c in stages.columns:
            stages[c] = stages[c].apply(parse_date)
    cases["Practice Area"] = cases["Practice Area"].astype(str).str.strip()
    cases["Case Stage"] = cases["Case Stage"].astype(str).str.strip()
    stages["Case Stage"] = stages["Case Stage"].astype(str).str.strip()
    stages["Case Number"] = stages["Case Number"].astype(str)

    hoje = datetime.now().date()

    # Função para somar dias em 'USCIS Pending Decision' por case
    def uscis_pending_days(case_no):
        sub = stages[stages["Case Number"].astype(str)==str(case_no)]
        if sub.empty: return 0
        mask = sub["Case Stage"].str.upper().str.contains("USCIS PENDING DECISION")
        sub = sub[mask].copy()
        if sub.empty: return 0
        sub["sd"] = sub["Start Date"].apply(lambda x: x or hoje)
        sub["ed"] = sub["End Date"].apply(lambda x: x or hoje)
        sub["d"] = (sub["ed"] - sub["sd"]).apply(lambda x: x.days if pd.notna(x) and x.days>=0 else 0)
        return int(sub["d"].sum())

    rows = []
    for _, r in cases.iterrows():
        case_no = r.get("Case Number")
        area    = r.get("Practice Area") or "(sem área)"
        od      = r.get("Open Date")
        cd      = r.get("Closed Date")
        sol     = r.get("Statute of Limitations Date")
        endref  = cd or hoje
        if not od:
            continue
        total_days = (endref - od).days if endref >= od else 0
        upd_days = uscis_pending_days(case_no)
        adj_days = max(0, total_days - upd_days)

        # SOL overrun (quanto passou do SOL até endref)
        sol_over = 0
        if sol:
            sol_over = max(0, (endref - sol).days)

        rows.append({"Practice Area": area, "Case Number": case_no, "AdjCompletionDays": adj_days, "SOL_OverrunDays": sol_over})

    if not rows:
        st.info("Sem dados suficientes para estimar tempos médios (verifique colunas e histórico de estágios).")
    else:
        est = pd.DataFrame(rows)
        resumo = est.groupby("Practice Area").agg(
            casos=("Case Number","count"),
            media_tempo_conclusao=("AdjCompletionDays","mean"),
            media_sol_ultrapasso=("SOL_OverrunDays","mean"),
            pct_ultrapassados=("SOL_OverrunDays", lambda s: 100.0* (s.gt(0).sum()/max(len(s),1)))
        ).round(1).sort_values("media_tempo_conclusao", ascending=False).reset_index()

        st.dataframe(resumo, use_container_width=True)

        # Gráfico: média de conclusão (ajustada) por área
        fig, ax = plt.subplots(figsize=(10, max(3, 0.5*len(resumo))))
        ax.barh(resumo["Practice Area"], resumo["media_tempo_conclusao"], color=ACCENT)
        ax.invert_yaxis()
        ax.set_xlabel("Dias (média ajustada)")
        ax.set_title("Tempo médio de conclusão (excl. USCIS Pending Decision) por Área")
        fig.patch.set_facecolor(BG_SOFT); ax.set_facecolor("#0B2C21")
        st.pyplot(fig, clear_figure=True)

        # Downloads
        st.download_button(
            "⬇️ Baixar estimativas (CSV)",
            data=df_to_csv_bytes(resumo, include_index=False),
            file_name="estimativas_por_area.csv",
            mime="text/csv"
        )
else:
    st.caption("Para estimar tempos médios e SOL: carregue **Casos** (com Open/Closed/SOL) e **Histórico de Estágios**.")
