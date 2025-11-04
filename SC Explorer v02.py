# social_contract_explorer_final_app_v153_theme_mapper.py
# Social Contract Explorer — v1.53 (Theme→Domain Admin Mapper)
# Author: CO3 Database
# Date: 2025-11-04
#
# v1.53 changes
# - Admin mode to manage attribution of Themes to SC Domains.
# - Reads/writes a 'themes' sheet: columns ['theme','domain_id'] (domain_id in 1..6).
# - Interactive mapping table with Selectbox; export mapping CSV and updated workbook.
# - Mapping overrides any text-based domain normalization in the pipeline.
#
# Place your Excel next to this file and name it 'unified_workbook.xlsx' for autoload.

import io
import re
from io import BytesIO
from pathlib import Path
from typing import List, Dict, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ---------- Config ----------
st.set_page_config(page_title="Social Contract Explorer", page_icon="🧭", layout="wide")

def _get_query_params():
    try:
        return st.query_params  # type: ignore[attr-defined]
    except Exception:
        try:
            return st.experimental_get_query_params()
        except Exception:
            return {}

params = _get_query_params()
EMBED = str(params.get("embed", "0")).strip().lower() in {"1", "true", "yes"}
ADMIN = str(params.get("admin", "0")).strip().lower() in {"1", "true", "yes"}

# Name of the bundled workbook to autoload
DEFAULT_WORKBOOK = "unified_workbook.xlsx"

# Base style
BASE_STYLE = """
<style>
.block-container{padding-top:1.0rem;padding-bottom:1.0rem;}
.stAlert{border-radius:12px;}
.metric-card{border:1px solid #E5E7EB;border-radius:14px;padding:14px;}
.caption{color:#6B7280;font-size:.9rem;}
.small{font-size:12px;color:#6B7280;}
</style>
"""
st.markdown(BASE_STYLE, unsafe_allow_html=True)

if EMBED:
    EMBED_STYLE = """
    <style>
    header {visibility: hidden;}
    [data-testid="stToolbar"] {display: none !important;}
    footer {visibility: hidden;}
    .block-container{padding-top:.4rem;padding-bottom:.4rem;}
    section[data-testid="stSidebar"] {border-right: 0;}
    </style>
    """
    st.markdown(EMBED_STYLE, unsafe_allow_html=True)

st.title("🧭 Social Contract Explorer — v1.53")
st.caption("Now with Theme → SC Domain admin mapping.")

DOMAINS_MAP = {1:"Legitimacy",2:"Fairness",3:"Citizenship",4:"Social Cohesion",5:"Citizen–State Relationship",6:"Resilience"}
DOMAIN_SYNONYMS = {"legitimacy":"Legitimacy","fairness":"Fairness","citizenship":"Citizenship","social cohesion":"Social Cohesion","citizen-state":"Citizen–State Relationship","citizen – state relationship":"Citizen–State Relationship","citizen-state relationship":"Citizen–State Relationship","citizen–state relationship":"Citizen–State Relationship","resilience":"Resilience","meşruiyet":"Legitimacy","adalet":"Fairness","yurttaşlık":"Citizenship","toplumsal uyum":"Social Cohesion","vatandaş-devlet ilişkisi":"Citizen–State Relationship","dayanıklılık":"Resilience"}
WAVE_GROUPS = [("1981–1984 (EVS/WVS1)", (1981, 1984), "EVS/WVS1"),("1989–1993 (EVS/WVS2)", (1989, 1993), "EVS/WVS2"),("1994–1998 (WVS3)", (1994, 1998), "WVS3"),("1999–2004 (EVS3/WVS4)", (1999, 2004), "EVS3/WVS4"),("2005–2010 (EVS4/WVS5)", (2005, 2010), "EVS4/WVS5"),("2010–2014 (WVS6)", (2010, 2014), "WVS6"),("2017–2022 (EVS5/WVS7)", (2017, 2022), "EVS5/WVS7")]
WAVE_LABELS = [w[0] for w in WAVE_GROUPS]

@st.cache_data(show_spinner=False)
def to_group_label(v) -> str:
    try:
        if isinstance(v, str):
            if any(v.startswith(lbl[:4]) for lbl, _, _ in WAVE_GROUPS): return v
            for label, (a, b), short in WAVE_GROUPS:
                if short in v or label in v: return label
            m = re.search(r"(19\\d{2}|20\\d{2})", v)
            if m:
                y = int(m.group(1))
                for label, (a, b), _ in WAVE_GROUPS:
                    if a <= y <= b: return label
        y = int(v)
        for label, (a, b), _ in WAVE_GROUPS:
            if a <= y <= b: return label
    except Exception:
        pass
    return "Unknown"

def normalize_domain_text(s: str) -> str:
    if s is None or (isinstance(s, float) and np.isnan(s)): return "Unassigned"
    t = str(s).strip().replace("_"," ").replace("–","-").lower()
    t = re.sub(r"\\s+"," ", t)
    return DOMAIN_SYNONYMS.get(t, t.title())

@st.cache_data(show_spinner=False)
def read_workbook(pathlike) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(pathlike)
    data = {name: pd.read_excel(pathlike, sheet_name=name) for name in xls.sheet_names}
    return data

def build_long(df_dict: Dict[str, pd.DataFrame], theme_mapping: pd.DataFrame | None) -> pd.DataFrame:
    need = {"surveys","surveys_questions","questions_names"}
    if not need.issubset(df_dict.keys()):
        raise ValueError("Unified sheets not found (need: surveys, surveys_questions, questions_names).")
    surveys = df_dict["surveys"]
    sq = df_dict["surveys_questions"]
    qn = df_dict["questions_names"][["question_code","question_name"]].drop_duplicates()

    m1 = sq.merge(surveys, on="survey_id", how="left")
    m2 = m1.merge(qn, on="question_code", how="left")
    long_df = m2.rename(columns={"question_name":"question_label"}).copy()
    long_df["available"] = 1

    # Theme column fallback
    if "theme" not in long_df.columns:
        long_df["theme"] = "All"

    # Apply explicit theme→domain mapping if provided
    if theme_mapping is not None and not theme_mapping.empty:
        tm = theme_mapping.copy()
        tm["theme"] = tm["theme"].astype(str)
        tm["domain_id"] = pd.to_numeric(tm["domain_id"], errors="coerce").astype("Int64")
        tm["domain"] = tm["domain_id"].map(DOMAINS_MAP).fillna("Unassigned")
        long_df = long_df.merge(tm[["theme","domain"]], on="theme", how="left", suffixes=("","_mapped"))
        # Prefer mapped domain if present; else fall back to theme_group_id or text
        long_df["domain"] = long_df["domain"].where(long_df["domain"].notna(), None)
    else:
        long_df["domain"] = None

    # Fill domain using theme_group_id if still missing
    if "theme_group_id" in long_df.columns:
        mask = long_df["domain"].isna()
        if mask.any():
            def map_dom(x):
                try: return DOMAINS_MAP.get(int(x), "Unassigned") if pd.notna(x) else "Unassigned"
                except Exception: return "Unassigned"
            long_df.loc[mask, "domain"] = long_df.loc[mask, "theme_group_id"].apply(map_dom)

    # Fallback to text normalization if still missing
    long_df["domain"] = long_df["domain"].fillna(long_df.get("theme","Unassigned").apply(normalize_domain_text))

    # Wave
    long_df["wave"] = long_df["year"].apply(to_group_label)

    out = long_df[["country","wave","domain","theme","question_code","question_label","available"]]\
        .dropna(subset=["country","wave","question_code"])
    return out

def extract_theme_mapping(df_dict: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    # If workbook has a 'themes' sheet with columns ['theme','domain_id'], use it.
    if "themes" in df_dict:
        t = df_dict["themes"].copy()
        cols = {c.lower(): c for c in t.columns}
        if "theme" in cols and "domain_id" in cols:
            t = t.rename(columns={cols["theme"]:"theme", cols["domain_id"]:"domain_id"})
            t = t[["theme","domain_id"]]
            return t
    # Else, infer from data (unique themes, unassigned domain_id)
    if {"surveys","surveys_questions"}.issubset(df_dict.keys()):
        sq = df_dict["surveys_questions"]
        if "theme" in sq.columns:
            themes = pd.DataFrame({"theme": sorted(sq["theme"].dropna().astype(str).unique().tolist())})
        else:
            themes = pd.DataFrame({"theme": ["All"]})
    else:
        themes = pd.DataFrame({"theme": ["All"]})
    themes["domain_id"] = pd.Series([pd.NA] * len(themes), dtype="Int64")
    return themes

def export_updated_workbook(df_dict: Dict[str, pd.DataFrame], theme_mapping: pd.DataFrame) -> bytes:
    # Write back all original sheets + 'themes' mapping sheet
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, df in df_dict.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
        # Ensure themes sheet exists and clean
        t = theme_mapping.copy()
        t = t.assign(domain_name=t["domain_id"].map(DOMAINS_MAP))
        t.to_excel(writer, sheet_name="themes", index=False)
    bio.seek(0)
    return bio.read()

with st.sidebar:
    st.header("Data & Admin")
    # Autoload workbook
    default_path = Path(__file__).parent / DEFAULT_WORKBOOK
    if default_path.exists():
        workbook_path = str(default_path)
        st.success(f"Auto-loaded bundled workbook: {DEFAULT_WORKBOOK}")
        override = st.file_uploader("📄 (Optional) Upload a different unified workbook", type=["xlsx", "xls"], help="Overrides the bundled file for this session.")
        if override is not None:
            workbook_path = override
    else:
        workbook_path = st.file_uploader("📄 Upload unified workbook (Excel)", type=["xlsx", "xls"])
        if not workbook_path:
            st.stop()

    admin_toggle = st.toggle("Admin mode (Theme→Domain mapper)", value=ADMIN)
    st.caption("Use admin mode to re-assign themes to the 6 SC domains.")

    if not EMBED:
        st.markdown("---")
        st.subheader("🔗 Embed")
        example_url = "https://your-org-your-app.streamlit.app/?embed=1"
        iframe = f'<iframe src="{example_url}" width="100%" height="900" style="border:0;" loading="lazy" referrerpolicy="no-referrer-when-downgrade" allowfullscreen></iframe>'
        st.code(iframe, language="html")

# Read workbook into memory
try:
    df_dict = read_workbook(workbook_path)
except Exception as e:
    st.error(f"Failed to read workbook: {e}")
    st.stop()

# Load or infer theme mapping
theme_map_df = extract_theme_mapping(df_dict)

if admin_toggle:
    st.subheader("Admin — Theme → SC Domain mapper")
    st.caption("Assign each theme to one of the 6 canonical SC domains. Changes apply immediately in the analysis below and can be exported as a new workbook (adds a 'themes' sheet).")

    # Add computed domain name for display
    tm = theme_map_df.copy()
    tm["domain_id"] = pd.to_numeric(tm["domain_id"], errors="coerce").astype("Int64")
    tm["domain_name"] = tm["domain_id"].map(DOMAINS_MAP)

    # Use data_editor with a SelectboxColumn for domain_id
    domain_options = list(DOMAINS_MAP.keys())
    # Prepare column configs
    col_config = {
        "theme": st.column_config.TextColumn("Theme", disabled=True),
        "domain_id": st.column_config.SelectboxColumn(
            "Domain (id)",
            options=domain_options,
            required=False,
            help="Select 1..6",
        ),
        "domain_name": st.column_config.TextColumn("Domain (name)", disabled=True)
    }

    edited = st.data_editor(
        tm,
        hide_index=True,
        use_container_width=True,
        column_config=col_config,
        num_rows="fixed",
        key="theme_editor",
    )

    # Reflect edits back
    theme_map_df = edited[["theme","domain_id"]].copy()

    # Download mapping CSV
    csv_buf = io.StringIO()
    theme_map_df.to_csv(csv_buf, index=False)
    st.download_button("⬇️ Download mapping (CSV)", csv_buf.getvalue(), "themes_mapping.csv", "text/csv")

    # Download updated workbook containing the new 'themes' sheet
    try:
        bytes_xlsx = export_updated_workbook(df_dict, theme_map_df)
        st.download_button("⬇️ Download updated workbook (XLSX)", data=bytes_xlsx, file_name="unified_workbook_updated.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.warning(f"Couldn't export workbook: {e}")

# Build long table using mapping (if any)
try:
    df_long = build_long(df_dict, theme_map_df)
except Exception as e:
    st.error(f"Failed to parse the unified workbook: {e}")
    st.stop()

# Metrics and analysis UI
colA, colB, colC, colD = st.columns(4)
with colA: st.metric("Countries", f"{df_long['country'].nunique():,}")
with colB: st.metric("Waves (grouped)", f"{df_long['wave'].nunique():,}")
with colC: st.metric("Domains", f"{df_long['domain'].nunique():,}")
with colD: st.metric("Questions", f"{df_long['question_code'].nunique():,}")
st.info("**Domain source:** 'themes' sheet mapping (if present) → theme_group_id → text normalization")

st.divider()
st.subheader("Stage 2 — Choose Domains")
canonical_order = [DOMAINS_MAP[i] for i in sorted(DOMAINS_MAP.keys())]
mapping = {}
for (dom, th), g in df_long.groupby(["domain", "theme"]):
    mapping.setdefault(str(dom), {}).setdefault(str(th), [])
    pairs = sorted(set(zip(g["question_code"].astype(str), g["question_label"].astype(str))))
    mapping[str(dom)][str(th)] = pairs

existing = [d for d in canonical_order if d in mapping]
chosen_domains = st.multiselect("Select one or more domains:", options=existing or canonical_order, default=existing[:1] or canonical_order[:1])
if not chosen_domains: st.stop()

st.subheader("Stage 3 — Themes & Questions")
selected_questions: List[str] = []
for dom in chosen_domains:
    st.markdown(f"### **{dom}**")
    themes = mapping.get(dom, {})
    cols = st.columns(2, gap="large")
    i = 0
    for th, qlist in themes.items():
        with cols[i % 2]:
            with st.container(border=True):
                st.markdown(f"**Theme:** {th}")
                for code, label in qlist:
                    if st.checkbox(f"{code} — {label}", key=f"sel_{dom}_{th}_{code}"):
                        selected_questions.append(str(code))
        i += 1
if not selected_questions:
    st.warning("Select at least one question to proceed."); st.stop()

st.subheader("Stage 4 — Countries & Waves")
available_countries = sorted(df_long["country"].dropna().astype(str).unique().tolist())
available_waves = [w for w in WAVE_LABELS if w in df_long["wave"].unique()]
c1, c2 = st.columns([2,1])
with c1: chosen_countries = st.multiselect("Countries", options=available_countries, default=available_countries[:20])
with c2: chosen_waves = st.multiselect("Waves (grouped)", options=available_waves, default=available_waves)

logic_mode = st.radio("Coverage logic", ["Any selected question present (OR)","All selected questions present (AND)","Share of selected questions present (0–1)"], index=0, horizontal=True if EMBED else False)

st.subheader("Stage 5 — Coverage Heatmap")
subset = df_long[(df_long["country"].isin(chosen_countries)) & (df_long["wave"].isin(chosen_waves)) & (df_long["question_code"].astype(str).isin([str(q) for q in selected_questions]))].copy()
if subset.empty:
    st.warning("No data after filtering. Try adding more countries, waves, or questions."); st.stop()

pivot = subset.pivot_table(index="country", columns=["wave","question_code"], values="available", aggfunc="max", fill_value=0)

def agg_vals(sub_mat: pd.DataFrame) -> np.ndarray:
    arr = sub_mat.to_numpy()
    if logic_mode.startswith("Any"): return (arr.max(axis=1) > 0).astype(float)
    elif logic_mode.startswith("All"): return (arr.min(axis=1) > 0).astype(float)
    else: return arr.mean(axis=1)

wave_vals, wave_hovers = {}, {}
present_idx = None
for w in chosen_waves:
    qcols = [c for c in pivot.columns if isinstance(c, tuple) and c[0] == w]
    if not qcols: continue
    sub_mat = pivot[qcols]
    vals = agg_vals(sub_mat)
    wave_vals[w] = vals
    present_idx = sub_mat.index
    present = (sub_mat == 1)
    labels_map = dict(subset[subset["wave"] == w][["question_code","question_label"]].drop_duplicates().values)
    h = []
    for i, ctry in enumerate(present.index.tolist()):
        present_qs = [qc[1] for j, qc in enumerate(qcols) if present.iloc[i, j] == 1]
        if not present_qs: h.append(f"{ctry} | {w}\\nNo questions available")
        else:
            items = [f"• {q} — {labels_map.get(q, q)}" for q in present_qs]
            h.append(f"{ctry} | {w}\\n" + "\\n".join(items))
    wave_hovers[w] = h

heat_df = pd.DataFrame(wave_vals)
if present_idx is not None: heat_df.index = present_idx
heat_df = heat_df.reindex(index=sorted(heat_df.index.tolist()))
heat_df = heat_df[[w for w in chosen_waves if w in heat_df.columns]]

hover_text = pd.DataFrame({w: wave_hovers[w] for w in heat_df.columns}, index=heat_df.index)

fig = px.imshow(heat_df.values, x=heat_df.columns.tolist(), y=heat_df.index.tolist(), aspect="auto", color_continuous_scale=[[0,"#F3F4F6"],[1,"#1F77B4"]] if not logic_mode.startswith("Share") else None, origin="upper")
if logic_mode.startswith("Share"): fig.update_traces(zmin=0, zmax=1, colorbar_title="Share")
else: fig.update_traces(zmin=0, zmax=1, showscale=False)
fig.update_traces(hovertemplate="%{customdata}", customdata=hover_text.values)
fig.update_layout(margin=dict(l=10,r=10,t=10), xaxis_title="Waves (grouped)", yaxis_title="Countries")
st.plotly_chart(fig, use_container_width=True, theme="streamlit")

st.subheader("Domain coverage summary")
total_cells = max(1, len(chosen_countries) * len(chosen_waves))
summary_rows = []
sel_set = set(map(str, selected_questions))
for dom in chosen_domains:
    dom_qs = [code for th, qs in mapping.get(dom, {}).items() for code, _ in qs if str(code) in sel_set]
    if not dom_qs:
        summary_rows.append({"domain": dom, "selected_questions": 0, "covered_cells": 0, "coverage_share": 0.0}); continue
    sub_dom = df_long[(df_long["country"].isin(chosen_countries)) & (df_long["wave"].isin(chosen_waves)) & (df_long["question_code"].astype(str).isin([str(q) for q in dom_qs]))]
    if sub_dom.empty:
        summary_rows.append({"domain": dom, "selected_questions": len(set(dom_qs)), "covered_cells": 0, "coverage_share": 0.0}); continue
    pv = sub_dom.pivot_table(index="country", columns=["wave","question_code"], values="available", aggfunc="max", fill_value=0)
    covered = 0
    for w in chosen_waves:
        qcols = [c for c in pv.columns if isinstance(c, tuple) and c[0] == w]
        if not qcols: continue
        vals = agg_vals(pv[qcols])
        covered += int(vals.sum())
    summary_rows.append({"domain": dom, "selected_questions": len(set(dom_qs)), "covered_cells": covered, "coverage_share": round(covered / total_cells, 3)})

summary_df = pd.DataFrame(summary_rows).sort_values(["coverage_share","domain"], ascending=[False, True])
st.dataframe(summary_df, use_container_width=True, hide_index=True)

st.subheader("Map — coverage by country (selected wave)")
if chosen_waves:
    map_wave = st.selectbox("Wave to map:", options=chosen_waves, index=0)
    if map_wave in heat_df.columns:
        map_vals = heat_df[map_wave].reset_index(); map_vals.columns = ["country","value"]
        fig_map = px.choropleth(map_vals, locations="country", locationmode="country names", color="value", color_continuous_scale=[[0,"#F3F4F6"],[1,"#1F77B4"]] if not logic_mode.startswith("Share") else "Blues", range_color=(0,1), title=None)
        fig_map.update_layout(margin=dict(l=0,r=0,t=0,b=0), height=420)
        st.plotly_chart(fig_map, use_container_width=True, theme="streamlit")
    else:
        st.info("No data for the selected wave to map.")

st.markdown("#### Downloads")
matrix_csv = heat_df.copy()
if not logic_mode.startswith("Share"): matrix_csv = (matrix_csv > 0).astype(int)
buf = io.StringIO(); matrix_csv.to_csv(buf); st.download_button("⬇️ Coverage matrix (CSV)", buf.getvalue(), "coverage_matrix.csv", "text/csv")
buf2 = io.StringIO(); subset.sort_values(["country","wave","domain","theme","question_code"]).to_csv(buf2, index=False); st.download_button("⬇️ Filtered availability (CSV)", buf2.getvalue(), "filtered_availability.csv", "text/csv")
