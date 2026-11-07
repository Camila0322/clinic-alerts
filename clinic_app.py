import streamlit as st
import pandas as pd
import re
import io
import openpyxl
import plotly.graph_objects as go
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime

# ============================================================================
# 1. PAGE CONFIG
# ============================================================================
st.set_page_config(
    page_title="Clinic Alert Extraction | USYD",
    page_icon="🚨",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================================
# 2. THEME (matches the main AST surveillance app)
# ============================================================================
BLUE_900 = "#0a2540"; BLUE_800 = "#103a6b"; BLUE_700 = "#14539a"
BLUE_600 = "#1d6fd6"; BLUE_500 = "#2f88e6"; BLUE_300 = "#93c2f4"
BLUE_50 = "#eef5fd"; INK = "#0e1c2b"; MUTED = "#5d6e80"
LINE = "#e2eaf3"; CANVAS = "#eaf0f7"; WHITE = "#ffffff"
RED = "#d63a4a"; AMBER = "#e0a400"; GREEN = "#1aa260"

ALERT_COLORS = {
    "Staphylococcus (OX/FOX)": RED,
    "E. coli (AMC/ENR)": "#1d6fd6",
    "Salmonella": AMBER,
}

st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
    html, body, [class*="css"], .stApp {{
        font-family:'Inter',-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
        color:{INK}; font-size:16px;
    }}
    .stApp {{ background-color:{CANVAS}; }}
    .block-container {{ padding-top:1.4rem; max-width:1340px; }}
    .amr-hero {{
        position:relative; overflow:hidden;
        background:linear-gradient(125deg,{BLUE_900} 0%,{BLUE_700} 55%,{BLUE_600} 100%);
        border-radius:18px; padding:2rem 2.2rem; margin-bottom:1.6rem;
        box-shadow:0 14px 34px rgba(10,37,64,0.28);
    }}
    .amr-hero::after {{
        content:""; position:absolute; top:-40%; right:-8%;
        width:360px; height:360px; border-radius:50%;
        background:radial-gradient(circle,rgba(255,255,255,0.12) 0%,rgba(255,255,255,0) 70%);
    }}
    .amr-hero h1 {{ color:{WHITE}; font-size:2.25rem; font-weight:800; margin:0;
        letter-spacing:-0.025em; line-height:1.12; }}
    .amr-hero p {{ color:#d3e4fb; font-size:1.05rem; margin:0.5rem 0 0 0; font-weight:500; }}
    .amr-eyebrow {{ display:inline-block; color:#bcd6f7; font-size:0.78rem; font-weight:700;
        letter-spacing:0.14em; text-transform:uppercase; margin-bottom:0.5rem; }}
    .sec {{ display:flex; align-items:center; gap:0.7rem; font-size:1.5rem; font-weight:800;
        color:{BLUE_800}; letter-spacing:-0.015em; margin:0.2rem 0 1rem 0; }}
    .sec-bar {{ width:6px; height:26px; border-radius:6px;
        background:linear-gradient(180deg,{BLUE_500},{BLUE_700}); display:inline-block; }}
    .sec-note {{ color:{MUTED}; font-size:1rem; margin:-0.6rem 0 1.1rem 1.1rem; }}
    [data-testid="stVerticalBlockBorderWrapper"] {{
        background:{WHITE}; border:1px solid {LINE} !important; border-radius:16px;
        box-shadow:0 4px 18px rgba(14,28,43,0.06); padding:1.2rem 1.4rem; }}
    [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stVerticalBlock"] {{
        border:none !important; background:transparent !important;
        box-shadow:none !important; padding:0 !important; }}
    .stButton>button {{ width:100%; border-radius:11px; height:3.2em; border:none;
        background:linear-gradient(135deg,{BLUE_600} 0%,{BLUE_500} 100%);
        color:{WHITE}; font-weight:700; font-size:1.02rem;
        box-shadow:0 6px 18px rgba(29,111,214,0.32); transition:all 0.16s ease; }}
    .stButton>button:hover {{ background:linear-gradient(135deg,{BLUE_700} 0%,{BLUE_600} 100%);
        box-shadow:0 9px 24px rgba(29,111,214,0.42); transform:translateY(-1px); color:{WHITE}; }}
    .stDownloadButton>button {{ border-radius:11px; background:{WHITE}; color:{BLUE_700};
        border:1.6px solid {BLUE_500}; font-weight:700; height:3em; }}
    .stDownloadButton>button:hover {{ background:{BLUE_600}; color:{WHITE}; border-color:{BLUE_600}; }}
    [data-testid="stSidebar"] {{ background:{WHITE}; border-right:1px solid {LINE}; }}
    [data-testid="stMetric"] {{
        background:linear-gradient(180deg,{WHITE} 0%,{BLUE_50} 130%);
        border:1px solid {LINE}; border-left:5px solid {RED};
        border-radius:14px; padding:1.25rem 1.45rem; box-shadow:0 3px 12px rgba(14,28,43,0.05); }}
    [data-testid="stMetricValue"] {{ color:{BLUE_700}; font-weight:900; font-size:2.6rem; }}
    [data-testid="stMetricLabel"] {{ color:{MUTED}; font-weight:700; font-size:1.1rem; }}
    [data-testid="stMetricLabel"] p {{ font-size:1.1rem; }}
    [data-testid="stFileUploader"] {{ background:transparent; border:none; padding:0; }}
    [data-testid="stFileUploaderDropzone"] {{
        background:{BLUE_50}; border:1.4px dashed {BLUE_300}; border-radius:11px;
        padding:0.5rem 0.9rem; min-height:auto; }}
    [data-testid="stDataFrame"] {{ border-radius:10px; overflow:hidden; border:none; }}
    .legend-chip {{ display:inline-block; width:14px; height:14px; border-radius:4px;
        margin-right:8px; vertical-align:middle; }}
    hr {{ border-color:{LINE}; }}
    </style>
""", unsafe_allow_html=True)


def section(title, note=None):
    st.markdown(f"<div class='sec'><span class='sec-bar'></span>{title}</div>", unsafe_allow_html=True)
    if note:
        st.markdown(f"<div class='sec-note'>{note}</div>", unsafe_allow_html=True)


def style_fig(fig, height, bottom=40):
    fig.update_layout(height=height, template="simple_white", plot_bgcolor="white",
        paper_bgcolor="rgba(0,0,0,0)", font=dict(color=INK, size=15, family="Inter"),
        margin=dict(b=bottom, t=24, l=4, r=4),
        hoverlabel=dict(font=dict(size=16, family="Inter"), bgcolor="white", bordercolor=LINE))
    fig.update_xaxes(showline=True, linewidth=1.4, linecolor=LINE, tickfont=dict(size=14))
    fig.update_yaxes(showline=True, linewidth=1.4, linecolor=LINE, gridcolor=LINE, tickfont=dict(size=14))
    return fig


# ============================================================================
# 3. AST SHEET CONFIG  (same layout as the main app)
# ============================================================================
AST_FIRST_ABX_COL = 13
AST_CP_COL = 1
AST_DATE_COL = 0
AST_CLINIC_COL = 3
AST_ISOLATE_COL = 10
AST_SKIP_SHEETS = {"MALDI ID NAME LIST"}

CANON_ABBREVS = [
    "AM10", "AMC30", "CZN 30", "CL30", "CAZ30", "CVN30", "DO30", "SXT", "ENR5",
    "MAR", "GM 10", "C30", "P10", "F/M300", "N30", "OX1", "FOX30", "CC2", "E15",
    "TIC75", "TIM85", "S5", "CN 120", "RD5", "FA10", "AN30", "IPM10", "VA30",
]

# Trigger antibiotics for the alert rules.
TRIGGER_ABX = ["AMC30", "ENR5", "OX1", "FOX30"]


def meas_col(abx):
    return AST_FIRST_ABX_COL + 2 * CANON_ABBREVS.index(abx)


def interp_col(abx):
    return AST_FIRST_ABX_COL + 1 + 2 * CANON_ABBREVS.index(abx)


def normalize_cp(value):
    if value is None:
        return None
    m = re.search(r'(\d{2}-\d{4,5})', str(value))
    return m.group(1) if m else None


def resolve_cell(row, abx):
    """S / I / R / INTR / NA for one antibiotic, from the (measurement, interp) pair."""
    mi, si = meas_col(abx), interp_col(abx)
    meas = row[mi] if mi < len(row) else None
    sir = row[si] if si < len(row) else None
    if meas is not None and str(meas).strip().upper() == "INTR":
        return "INTR"
    if sir is None:
        return "NA"
    s = str(sir).strip().upper()
    if s in ("S", "I", "R"):
        return s
    if s in ("", "N/A", "NA"):
        return "NA"
    return s  # unexpected value, kept verbatim


def fmt_date(v):
    if v is None or v == "":
        return "NA"
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    try:
        dt = pd.to_datetime(v, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    return str(v)


# ============================================================================
# 4. ALERT EXTRACTION
# ============================================================================
# Rules (clinically significant findings to flag back to clinics):
#   1. Escherichia coli resistant to AMC (AMC30) or enrofloxacin (ENR5)
#   2. Any Staphylococcus resistant to oxacillin (OX1) or cefoxitin (FOX30)
#      (methicillin resistance marker: MRSA / MRSP)
#   3. Any Salmonella, regardless of susceptibility profile
def extract_alerts(file_object):
    wb = openpyxl.load_workbook(file_object, read_only=True, data_only=True)
    records = []
    for sheet_name in wb.sheetnames:
        if sheet_name in AST_SKIP_SHEETS:
            continue
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header = rows[0]
        if not header or len(header) <= AST_CP_COL or str(header[AST_CP_COL]).strip().upper() != "CP NUMBER":
            continue
        for row in rows[1:]:
            if len(row) <= AST_ISOLATE_COL or not row[AST_CP_COL] or not row[AST_ISOLATE_COL]:
                continue
            org = str(row[AST_ISOLATE_COL]).strip()
            genus = org.split()[0].lower() if org else ""
            cp = normalize_cp(row[AST_CP_COL])
            clinic = str(row[AST_CLINIC_COL]).strip() if len(row) > AST_CLINIC_COL and row[AST_CLINIC_COL] else "NA"
            date = fmt_date(row[AST_DATE_COL] if len(row) > AST_DATE_COL else None)
            results = {a: resolve_cell(row, a) for a in TRIGGER_ABX}

            alert_type, reason = None, None
            if genus == "escherichia" and "coli" in org.lower():
                hits = [a for a in ("AMC30", "ENR5") if results[a] == "R"]
                if hits:
                    alert_type = "E. coli (AMC/ENR)"
                    reason = "Resistant to " + " and ".join(hits)
            elif genus == "staphylococcus":
                hits = [a for a in ("OX1", "FOX30") if results[a] == "R"]
                if hits:
                    alert_type = "Staphylococcus (OX/FOX)"
                    reason = "Resistant to " + " and ".join(hits) + " (methicillin resistance)"
            elif genus == "salmonella":
                alert_type = "Salmonella"
                reason = "Salmonella isolate (any profile)"

            if alert_type:
                records.append({
                    "Year": sheet_name,
                    "Date Received": date,
                    "CP Number": f"CP {cp}" if cp else "NA",
                    "Clinic": clinic,
                    "Organism": org.capitalize(),
                    "Alert": alert_type,
                    "Reason": reason,
                    "AMC30": results["AMC30"],
                    "ENR5": results["ENR5"],
                    "OX1": results["OX1"],
                    "FOX30": results["FOX30"],
                })

    df = pd.DataFrame(records, columns=[
        "Year", "Date Received", "CP Number", "Clinic", "Organism",
        "Alert", "Reason", "AMC30", "ENR5", "OX1", "FOX30"])
    if not df.empty:
        df = df.drop_duplicates(subset=["CP Number", "Organism", "Alert", "Reason"]).reset_index(drop=True)
        df = df.sort_values(["Year", "Clinic", "CP Number"]).reset_index(drop=True)
    return df


def sir_css(v):
    if v == "R":
        return "background-color: #FFC7CE; font-weight:700"
    if v == "I":
        return "background-color: #FFEB9C"
    if v == "S":
        return "background-color: #C6EFCE"
    return ""


def build_excel(df):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Clinic Alerts"
    cols = list(df.columns)
    hfont = Font(bold=True, color="103A6B", size=11)
    hborder = Border(bottom=Side(style="medium", color="103A6B"))
    fills = {"R": "FFC7CE", "I": "FFEB9C", "S": "C6EFCE"}
    for j, name in enumerate(cols, 1):
        c = ws.cell(1, j, name)
        c.font = hfont
        c.border = hborder
        c.alignment = Alignment(horizontal="left", vertical="center")
    for i, (_, row) in enumerate(df.iterrows(), start=2):
        for j, name in enumerate(cols, 1):
            v = row[name]
            cell = ws.cell(i, j, "" if (v is None or (isinstance(v, float) and pd.isna(v))) else v)
            if name in TRIGGER_ABX and str(v).strip() in fills:
                cell.fill = PatternFill("solid", fgColor=fills[str(v).strip()])
    widths = {"Year": 8, "Date Received": 14, "CP Number": 12, "Clinic": 22,
              "Organism": 30, "Alert": 20, "Reason": 34}
    for j, name in enumerate(cols, 1):
        letter = ws.cell(1, j).column_letter
        ws.column_dimensions[letter].width = widths.get(name, 8)
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================================
# 5. SIDEBAR
# ============================================================================
with st.sidebar:
    st.markdown(f"<h3 style='color:{BLUE_800};font-size:1.3rem;font-weight:800;margin-bottom:0.4rem;'>What is flagged</h3>",
                unsafe_allow_html=True)
    st.markdown(
        f"<span class='legend-chip' style='background:{RED};'></span>**Staphylococcus** resistant to OX1 or FOX30<br>"
        f"<span style='color:{MUTED};font-size:0.85rem;'>methicillin resistance (MRSA / MRSP)</span><br><br>"
        f"<span class='legend-chip' style='background:{BLUE_600};'></span>**E. coli** resistant to AMC30 or ENR5<br><br>"
        f"<span class='legend-chip' style='background:{AMBER};'></span>**Salmonella** (any profile)",
        unsafe_allow_html=True)
    st.divider()
    st.markdown("**Antibiotic cells**")
    st.markdown(
        f"<span class='legend-chip' style='background:#C6EFCE;'></span>Susceptible (S)<br>"
        f"<span class='legend-chip' style='background:#FFEB9C;'></span>Intermediate (I)<br>"
        f"<span class='legend-chip' style='background:#FFC7CE;'></span>Resistant (R)",
        unsafe_allow_html=True)

# ============================================================================
# 6. MAIN
# ============================================================================
st.markdown(
    "<div class='amr-hero'>"
    "<span class='amr-eyebrow'>Sydney School of Veterinary Science</span>"
    "<h1>Clinic Alert Extraction</h1>"
    "<p>Flagging clinically significant isolates from the AST logging sheet</p>"
    "</div>",
    unsafe_allow_html=True)

with st.container(border=True):
    section("Upload the AST logging sheet", "Alerts are extracted from the AST logging workbook only. No reports are needed.")
    ast_file = st.file_uploader("AST LOGGING sheet(s) (Excel)", type=["xlsx"])
    run = st.button("Extract alerts")

if run:
    if not ast_file:
        st.error("Please upload the AST logging sheet.")
    else:
        with st.spinner("Scanning for flagged isolates..."):
            st.session_state["alerts"] = extract_alerts(ast_file)

if "alerts" in st.session_state:
    df = st.session_state["alerts"]

    if df.empty:
        st.warning("No isolates matched the alert rules in this sheet.")
    else:
        # ---- KPI strip ----
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total alerts", len(df))
        m2.metric("Staphylococcus (OX/FOX)", int((df["Alert"] == "Staphylococcus (OX/FOX)").sum()))
        m3.metric("E. coli (AMC/ENR)", int((df["Alert"] == "E. coli (AMC/ENR)").sum()))
        m4.metric("Salmonella", int((df["Alert"] == "Salmonella").sum()))
        st.write("")

        # ---- Alerts by clinic ----
        with st.container(border=True):
            section("Alerts by clinic", "Which clinics the flagged isolates came from.")
            clinic_counts = df["Clinic"].value_counts().sort_values(ascending=True)
            ys, xs = clinic_counts.index.tolist(), [int(v) for v in clinic_counts.values]
            fig = go.Figure(go.Bar(orientation="h", y=ys, x=xs, marker_color=BLUE_600,
                marker_line_width=0, text=xs, textposition="outside",
                textfont=dict(size=14, color=INK), cliponaxis=False,
                hovertemplate="<b>%{y}</b><br>Alerts: %{x}<extra></extra>"))
            style_fig(fig, max(280, 42 * len(ys)), bottom=30)
            fig.update_layout(xaxis_title="<b>Number of alerts</b>", yaxis_title="")
            fig.update_xaxes(range=[0, (max(xs) if xs else 1) * 1.2])
            st.plotly_chart(fig, use_container_width=True)

        st.write("")

        # ---- Alerts by type over time ----
        with st.container(border=True):
            section("Alerts over time", "Flagged isolates per year, split by alert type.")
            years = sorted(df["Year"].unique())
            figt = go.Figure()
            for atype, color in ALERT_COLORS.items():
                counts = [int(((df["Year"] == y) & (df["Alert"] == atype)).sum()) for y in years]
                figt.add_bar(x=years, y=counts, name=atype, marker_color=color,
                    hovertemplate="<b>%{x}</b><br>" + atype + ": %{y}<extra></extra>")
            style_fig(figt, 430, bottom=40)
            figt.update_layout(barmode="group", xaxis_title="<b>Year</b>",
                yaxis_title="<b>Number of alerts</b>",
                legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=13)))
            figt.update_xaxes(type="category")
            st.plotly_chart(figt, use_container_width=True)

        st.write("")

        # ---- Detailed flagged-isolate table ----
        with st.container(border=True):
            section("Flagged isolates", "Filter by clinic or alert type. Download sends the list to Excel.")
            f1, f2 = st.columns(2)
            with f1:
                clinics = ["All clinics"] + sorted(df["Clinic"].unique())
                pick_clinic = st.selectbox("Clinic", clinics)
            with f2:
                types = ["All types"] + sorted(df["Alert"].unique())
                pick_type = st.selectbox("Alert type", types)

            view = df.copy()
            if pick_clinic != "All clinics":
                view = view[view["Clinic"] == pick_clinic]
            if pick_type != "All types":
                view = view[view["Alert"] == pick_type]
            view = view.reset_index(drop=True)
            view.index = range(1, len(view) + 1)

            st.markdown(f"**{len(view)} isolate(s) shown**")
            st.dataframe(view.style.map(sir_css, subset=TRIGGER_ABX), use_container_width=True)

            aus_time = datetime.utcnow()
            fname = f"Clinic_Alerts_{aus_time.strftime('%Y%m%d')}.xlsx"
            st.download_button("Download alerts (Excel)", build_excel(view), fname,
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Upload the AST logging sheet and select Extract alerts to begin.")
