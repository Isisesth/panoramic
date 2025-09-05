import re
import io
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, date, timedelta
from docx import Document

# =========================
# CONFIG & TEMA
# =========================
st.set_page_config(page_title="USA4ALL • Panorama", layout="wide", page_icon="🗂️")

PRIMARY = "#0B3D2E"   # verde escuro
ACCENT  = "#9BE84F"   # verde-lima
BG_SOFT = "#0F4737"
LOGO_URL = "https://i.ytimg.com/vi/aWcON7jyz0I/hq720.jpg"

st.markdown(
    f"""
    <style>
    .stApp {{
        background: linear-gradient(180deg, {BG_SOFT} 0%, #07281F 100%);
        color: #ffffff;
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
    .stMetric-value, .stMetric-label {{ color: #ffffff !important; }}
    .stProgress > div > div > div > div {{ background-color: {ACCENT} !important; }}
    </style>
    """,
    unsafe_allow_html=True
)

# =========================
# HELPERS
# =========================
def parse_date(val, fallback=None):
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

def extract_stage_and_days(stage_text: str):
    """
    Retorna (stage_limpo, dias_em_parenteses:int).
    Ex.: "USCIS Pending Decision (23)" -> ("USCIS Pending Decision", 23)
    """
    if not isinstance(stage_text, str):
        return "", 0
    # pega o último número inteiro entre parênteses
    m = re.findall(r"\((\s*\d+\s*)\)", stage_text)
    days = int(m[-1].strip()) if m else 0
    # remove todo "(...)" do label
    stage_clean = re.sub(r"\s*\([^)]*\)\s*", "", stage_text).strip()
    return stage_clean, days

# Prazos SOL por área
SOL_PRAZO = {
    "FOIA": 30, "I-130": 30, "COS": 30, "B2-EXT": 30, "NPT": 60, "NVC": 60, "K1": 30,
    "WAIVER": 90, "EB2-NIW": 120, "EB1": 90, "E2": 90, "O1": 90, "EB4": 90,
    "AOS": 60, "ROC": 60, "NATZ": 30, "DA": 90, "PIP": 30, "I-90": 30, "COURT": 60,
    "DOM": 30, "SIJS": 30, "VAWA": 90, "T-VISA": 120, "U-VISA": 90, "I-918B": 30,
    "ASYLUM": 120, "PERM": 30, "EB3": 30
}

# =========================
# SIDEBAR
# =========================
with st.sidebar:
    st.image(LOGO_URL, caption="USA4ALL", use_column_width=True)
    mode = st.radio("Preenchimento", ["A partir de arquivo", "Manual"])

st.title("🗂️ Panorama de Casos — USA4ALL")

# =========================
# SESSION
# =========================
if "df_cases" not in st.session_state:
    st.session_state.df_cases = None
if "df_stages" not in st.session_state:
    st.session_state.df_stages = None

# =========================
# MAPEAMENTO
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
    for syn in synonyms:
        s = syn.strip().lower()
        for c, cl in norm.items():
            if cl == s:
                return c
    for syn in synonyms:
        s = syn.strip().lower()
        for c, cl in norm.items():
            if s in cl:
                return c
    return None

def mapping_ui(df, expected_dict, title):
    st.markdown(f"#### 🔎 Mapeamento — {title}")
    cols = list(df.columns)
    options = ["(não usar)"] + cols
    mapping = {}
    l, r = st.columns(2)
    keys = list(expected_dict.keys())
    half = len(keys)//2
    with l:
        for k in keys[:half]:
            sug = suggest_mapping(cols, expected_dict[k]) or "(não usar)"
            mapping[k] = st.selectbox(k, options, index=options.index(sug), key=f"map_{title}_{k}")
    with r:
        for k in keys[half:]:
            sug = suggest_mapping(cols, expected_dict[k]) or "(não usar)"
            mapping[k] = st.selectbox(k, options, index=options.index(sug), key=f"map_{title}_{k}")

    rename = {v: k for k, v in mapping.items() if v != "(não usar)"}
    df2 = df.rename(columns=rename).copy()
    # trim
    for c in df2.columns:
        if isinstance(c, str):
            df2[c] = df2[c].apply(lambda x: x.strip() if isinstance(x, str) else x)
    # datas
    for dc in ["Open Date","Closed Date","Statute of Limitations Date","Start Date","End Date"]:
        if dc in df2.columns:
            df2[dc] = pd.to_datetime(df2[dc], errors="coerce", dayfirst=True).dt.date
    st.success("✔️ Mapeamento aplicado.")
    return df2

# =========================
# UPLOADS
# =========================
df_cases = None
df_stages = None
selected_case = None

if mode == "A partir de arquivo":
    st.subheader("📂 Upload de Arquivos")
    up1 = st.file_uploader("Casos (CSV/XLS/XLSX)", type=["csv","xls","xlsx"], key="up_cases")
    up2 = st.file_uploader("Histórico de Estágios (opcional) — colunas: Case Number, Case Stage, Start Date, End Date",
                           type=["csv","xls","xlsx"], key="up_stages")

    if up1:
        try:
            raw_cases = pd.read_excel(up1) if up1.name.lower().endswith((".xls",".xlsx")) else pd.read_csv(up1)
            with st.expander("Ajustar colunas (Casos)"):
                df_cases = mapping_ui(raw_cases, CASES_FIELDS, "Casos")
            st.session_state.df_cases = df_cases
            st.success("✅ Casos carregados.")
            st.dataframe(df_cases.head(50), use_container_width=True)
        except Exception as e:
            st.error(f"Erro ao ler Casos: {e}")

    if up2:
        try:
            raw_stages = pd.read_excel(up2) if up2.name.lower().endswith((".xls",".xlsx")) else pd.read_csv(up2)
            with st.expander("Ajustar colunas (Histórico)"):
                df_stages = mapping_ui(raw_stages, STAGES_FIELDS, "Estágios")
            st.session_state.df_stages = df_stages
            st.success("✅ Histórico de Estágios carregado.")
            st.dataframe(df_stages.head(50), use_container_width=True)
        except Exception as e:
            st.error(f"Erro ao ler Estágios: {e}")

    if st.session_state.df_cases is not None:
        df_cases = st.session_state.df_cases
    if st.session_state.df_stages is not None:
        df_stages = st.session_state.df_stages

    # Seletor de cliente
    if df_cases is not None and "Case Number" in df_cases.columns:
        st.subheader("🔎 Selecione um cliente")
        options = df_cases["Case Number"].dropna().astype(str).unique().tolist()
        if options:
            selected_case = st.selectbox("Case Number", options)

    # ======= PAINEL DO CASE SELECIONADO =======
    if selected_case and df_cases is not None:
        row = df_cases[df_cases["Case Number"].astype(str) == str(selected_case)].iloc[0].to_dict()
        nome          = row.get("Case","")
        area          = row.get("Practice Area","")
        case_stage    = row.get("Case Stage","")
        open_date     = parse_date(row.get("Open Date"))
        closed_date   = parse_date(row.get("Closed Date"))
        sol_date      = parse_date(row.get("Statute of Limitations Date"))
        stage_clean, stage_days = extract_stage_and_days(case_stage)

        st.markdown(f"### 📌 {area or '—'} • **{nome or '—'}**  \n**Case Number:** {selected_case}")
        st.caption(f"Stage atual: **{stage_clean or '—'}**  |  dias no stage (via **()**): **{stage_days}**")

        # === TEMPO DO PROCESSO (NOVA REGRA)
        # base = hoje - Open Date
        hoje = datetime.now().date()
        if open_date:
            base_days = ( (closed_date or hoje) - open_date ).days
            if base_days < 0: base_days = 0
        else:
            base_days = 0

        # Se USCIS Pending Decision (no nome do stage), subtrai os dias entre parenteses
        if "USCIS PENDING DECISION" in (stage_clean or "").upper():
            tempo_processo = max(0, base_days - stage_days)
        else:
            tempo_processo = base_days

        # Progresso vs SOL (opcional, informativo)
        if open_date and sol_date:
            tot = (sol_date - open_date).days
            dec = (hoje - open_date).days
            rest = (sol_date - hoje).days
            perc = round((dec/tot)*100,2) if tot>0 else (100.0 if hoje>=sol_date else 0.0)
        else:
            tot=dec=rest=0; perc=0.0

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Tempo do processo (dias)", f"{tempo_processo}")
        c2.metric("Dias decorridos desde Open Date", f"{max(dec,0)}")
        c3.metric("Dias até SOL", f"{max(rest,0)}")
        c4.metric("Progresso do SOL", f"{perc:.2f}%")
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

        # --- GRÁFICO: duração por Case Stage (apenas ao selecionar cliente)
        st.subheader("⏱️ Duração por Case Stage — Cliente selecionado")
        if df_stages is not None and all(c in df_stages.columns for c in ["Case Number","Case Stage","Start Date","End Date"]):
            hist = df_stages[df_stages["Case Number"].astype(str)==str(selected_case)].copy()
            if not hist.empty:
                # calcula duração real Start→End; se End vazio, usa hoje
                def dur(r):
                    sd = parse_date(r.get("Start Date"))
                    ed = parse_date(r.get("End Date"), fallback=hoje)
                    if sd and ed and ed >= sd:
                        return (ed - sd).days
                    return 0
                hist["Dias"] = hist.apply(dur, axis=1)
                hist["Start Date"] = pd.to_datetime(hist["Start Date"], errors="coerce", dayfirst=True).dt.date
                hist["End Date"]   = pd.to_datetime(hist["End Date"], errors="coerce", dayfirst=True).dt.date
                hist = hist.sort_values("Start Date")

                fig_h = max(3, 0.5 * len(hist))
                fig, ax = plt.subplots(figsize=(10, fig_h))
                y = range(len(hist))
                ax.barh(y, hist["Dias"].tolist(), color=ACCENT)
                labels = [
                    f"{r['Case Stage']} ({fmt_date(r['Start Date'])} → {fmt_date(r['End Date'] or hoje)})"
                    for _, r in hist.iterrows()
                ]
                ax.set_yticks(list(y))
                ax.set_yticklabels(labels)
                ax.invert_yaxis()
                ax.set_xlabel("Dias")
                ax.set_title("Tempo em cada Stage")
                fig.patch.set_facecolor(BG_SOFT); ax.set_facecolor("#0B2C21")
                st.pyplot(fig, clear_figure=True)
            else:
                # fallback: usa o número entre parênteses do stage atual
                if stage_days > 0:
                    fig, ax = plt.subplots(figsize=(8, 2.8))
                    ax.barh([stage_clean], [stage_days], color=ACCENT)
                    ax.set_xlabel("Dias")
                    ax.set_title("Duração do stage atual (via '()')")
                    fig.patch.set_facecolor(BG_SOFT); ax.set_facecolor("#0B2C21")
                    st.pyplot(fig, clear_figure=True)
                else:
                    st.info("Sem histórico de estágios e sem valor '()' no stage atual.")
        else:
            # sem histórico: usa o '()' do stage atual
            if stage_days > 0:
                fig, ax = plt.subplots(figsize=(8, 2.8))
                ax.barh([stage_clean], [stage_days], color=ACCENT)
                ax.set_xlabel("Dias")
                ax.set_title("Duração do stage atual (via '()')")
                fig.patch.set_facecolor(BG_SOFT); ax.set_facecolor("#0B2C21")
                st.pyplot(fig, clear_figure=True)
            else:
                st.info("Envie o Histórico de Estágios para detalhar por estágio ou inclua dias entre '()' no Case Stage atual.")

# =========================
# OVERVIEW POR ÁREA (ATIVOS)
# =========================
st.subheader("🏢 Overview do Departamento — Casos ativos por Practice Area")
dfc = st.session_state.df_cases
if dfc is not None and not dfc.empty and "Practice Area" in dfc.columns and "Case Stage" in dfc.columns:
    x = dfc.copy()
    x["Practice Area"] = x["Practice Area"].astype(str).str.strip()
    def is_active(stage):
        s = str(stage).upper()
        return not(("APPROVED" in s) or ("DENIED" in s) or ("CLOSED" in s))
    x["Ativo"] = x["Case Stage"].apply(is_active)
    resumo = x.groupby("Practice Area")["Ativo"].sum().sort_values(ascending=False).reset_index(name="Casos Ativos")
    st.dataframe(resumo, use_container_width=True)

    fig, ax = plt.subplots(figsize=(10, max(3, 0.5*len(resumo))))
    ax.barh(resumo["Practice Area"], resumo["Casos Ativos"], color=ACCENT)
    ax.invert_yaxis()
    ax.set_xlabel("Casos ativos")
    ax.set_title("Ativos por Practice Area")
    fig.patch.set_facecolor(BG_SOFT); ax.set_facecolor("#0B2C21")
    st.pyplot(fig, clear_figure=True)
else:
    st.caption("Carregue o arquivo de **Casos** com colunas 'Practice Area' e 'Case Stage'.")

# =========================
# DIAS POR CASE STAGE (GERAL, via "()")
# =========================
st.subheader("📊 Dias por Case Stage (geral) — usando número entre '()' do campo Case Stage")
dfc2 = st.session_state.df_cases
if dfc2 is not None and not dfc2.empty and "Case Stage" in dfc2.columns:
    tmp = dfc2.copy()
    tmp["Stage Clean"], tmp["Stage Days"] = zip(*tmp["Case Stage"].apply(extract_stage_and_days))
    # considerar somente registros com Stage Clean não vazio e Stage Days > 0
    tmp = tmp[(tmp["Stage Clean"].astype(str).str.strip() != "") & (tmp["Stage Days"] > 0)]
    if tmp.empty:
        st.info("Nenhum valor entre '()' encontrado nos Case Stages.")
    else:
        stats = tmp.groupby("Stage Clean")["Stage Days"].agg(["count","mean","median","max"]).round(1).sort_values("mean", ascending=False).reset_index()
        stats = stats.rename(columns={"count":"#Casos", "mean":"Média (dias)", "median":"Mediana", "max":"Máx"})
        st.dataframe(stats, use_container_width=True)

        fig, ax = plt.subplots(figsize=(10, max(3, 0.45*len(stats))))
        ax.barh(stats["Stage Clean"], stats["Média (dias)"], color=ACCENT)
        ax.invert_yaxis()
        ax.set_xlabel("Média de dias (via '()')")
        ax.set_title("Média de dias por Case Stage (geral)")
        fig.patch.set_facecolor(BG_SOFT); ax.set_facecolor("#0B2C21")
        st.pyplot(fig, clear_figure=True)
else:
    st.caption("Carregue o arquivo de **Casos** com 'Case Stage' para calcular o gráfico de médias por stage.")

# =========================
# ESTIMATIVA: TEMPO MÉDIO DE CONCLUSÃO POR ÁREA
# Regra: total = (endref − Open Date), se stage tiver "USCIS Pending Decision (X)", subtrai X do total
# + SOL overrun médio e % de casos ultrapassados
# =========================
st.subheader("⏳ Estimativa — Tempo médio de conclusão por Practice Area (exclui 'USCIS Pending Decision' via '()') + SOL")

cases_est = st.session_state.df_cases
if cases_est is not None and not cases_est.empty and "Open Date" in cases_est.columns and "Case Stage" in cases_est.columns and "Practice Area" in cases_est.columns:
    c = cases_est.copy()
    c["Open Date"] = c["Open Date"].apply(parse_date)
    c["Closed Date"] = c["Closed Date"].apply(parse_date)
    c["Statute of Limitations Date"] = c["Statute of Limitations Date"].apply(parse_date)
    c["Practice Area"] = c["Practice Area"].astype(str).str.strip()

    # extrai info do Case Stage
    stage_clean_list = []
    stage_days_list  = []
    for s in c["Case Stage"].astype(str):
        sc, sd = extract_stage_and_days(s)
        stage_clean_list.append(sc)
        stage_days_list.append(sd)
    c["Stage Clean"] = stage_clean_list
    c["Stage Days"]  = stage_days_list

    hoje = datetime.now().date()
    rows = []
    for _, r in c.iterrows():
        od  = r.get("Open Date")
        cd  = r.get("Closed Date")
        sol = r.get("Statute of Limitations Date")
        area= r.get("Practice Area") or "(sem área)"
        sc  = (r.get("Stage Clean") or "").upper()
        sd  = int(r.get("Stage Days") or 0)

        if not od:
            continue
        endref = cd or hoje
        total_days = (endref - od).days
        if total_days < 0: total_days = 0

        # subtrai apenas se USCIS Pending Decision
        if "USCIS PENDING DECISION" in sc:
            adj = max(0, total_days - sd)
        else:
            adj = total_days

        # SOL overrun
        over = 0
        if sol:
            over = max(0, (endref - sol).days)

        rows.append({"Practice Area": area, "AdjCompletionDays": adj, "SOL_OverrunDays": over})

    if not rows:
        st.info("Sem dados suficientes (verifique Open Date / Case Stage / Practice Area).")
    else:
        est = pd.DataFrame(rows)
        resumo = est.groupby("Practice Area").agg(
            casos=("AdjCompletionDays","count"),
            media_tempo_conclusao=("AdjCompletionDays","mean"),
            media_sol_ultrapasso=("SOL_OverrunDays","mean"),
            pct_ultrapassados=("SOL_OverrunDays", lambda s: 100.0 * (s.gt(0).sum()/max(len(s),1)))
        ).round(1).sort_values("media_tempo_conclusao", ascending=False).reset_index()

        st.dataframe(resumo, use_container_width=True)

        fig, ax = plt.subplots(figsize=(10, max(3, 0.5*len(resumo))))
        ax.barh(resumo["Practice Area"], resumo["media_tempo_conclusao"], color=ACCENT)
        ax.invert_yaxis()
        ax.set_xlabel("Dias (média ajustada)")
        ax.set_title("Tempo médio de conclusão (ajustado) por Practice Area")
        fig.patch.set_facecolor(BG_SOFT); ax.set_facecolor("#0B2C21")
        st.pyplot(fig, clear_figure=True)

        st.download_button(
            "⬇️ Baixar estimativas por área (CSV)",
            data=df_to_csv_bytes(resumo, include_index=False),
            file_name="estimativas_conclusao_por_area.csv",
            mime="text/csv"
        )
else:
    st.caption("Carregue **Casos** com Open Date / Case Stage / Practice Area para a estimativa.")
