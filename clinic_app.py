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
# 2. THEME (University of Sydney ochre)
# ============================================================================
OCHRE_900 = "#5f1a0c"; OCHRE_800 = "#8c2a14"; OCHRE_700 = "#b0371c"
OCHRE_600 = "#e64626"; OCHRE_500 = "#ef6a4d"; OCHRE_300 = "#f4a98f"
OCHRE_50 = "#fcefe9"; INK = "#2e2d2c"; MUTED = "#6c6b6a"
LINE = "#ece3df"; CANVAS = "#f5efe9"; WHITE = "#ffffff"

# Alert-type colours (distinct, readable)
STAPH_C = "#b0371c"   # ochre-red
ECOLI_C = "#d99021"   # gold
SALM_C = "#4f6d7a"    # slate

ALERT_COLORS = {
    "Staphylococcus (OX/FOX)": STAPH_C,
    "E. coli (AMC/ENR)": ECOLI_C,
    "Salmonella": SALM_C,
}

LOGO_URI = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAABOQAAAHLCAYAAABoA0dYAAAWfmNhQlgAABZ+anVtYgAAAB5qdW1kYzJwYQARABCAAACqADibcQNjMnBhAAAAFlhqdW1iAAAAR2p1bWRjMm1hABEAEIAAAKoAOJtxA3VybjpjMnBhOmJkMTUxZWU2LTYzOGUtNGE0MC1iZDNhLWRlOGUzODI0NzU4YgAAAAOTanVtYgAAAClqdW1kYzJhcwARABCAAACqADibcQNjMnBhLmFzc2VydGlvbnMAAAAAuGp1bWIAAABEanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5pbmdyZWRpZW50LnYzAAAAABhjMnNoM5srGwgxT57cRUcx/1GXDAAAAGxjYm9yo2lkYzpmb3JtYXRpaW1hZ2UvcG5namluc3RhbmNlSUR4LHhtcDppaWQ6OGI4NjgzYzAtMTE2Zi00YzJhLThhODAtMDM4ZDNmNGRhYzA1bHJlbGF0aW9uc2hpcGhwYXJlbnRPZgAAAeJqdW1iAAAAQWp1bWRjYm9yABEAEIAAAKoAOJtxE2MycGEuYWN0aW9ucy52MgAAAAAYYzJzaMfCpIrX84OvcWh2PzGy1/oAAAGZY2JvcqJnYWN0aW9uc4KiZmFjdGlvbmtjMnBhLm9wZW5lZGpwYXJhbWV0ZXJzoWtpbmdyZWRpZW50c4GiY3VybHgtc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5pbmdyZWRpZW50LnYzZGhhc2hYIMpFNfLKO1Bxzc62QGnyJx/41fNkYAXD9h/5pXWZTas8pGZhY3Rpb254HWNvbS5hbnRocm9waWMuY2xhdWRlLnByb3ZpZGVkanBhcmFtZXRlcnOheB9jb20uYW50aHJvcGljLm9yaWdpbi1jb25maWRlbmNlZ3Vua25vd25rZGVzY3JpcHRpb254ZkNsYXVkZSBwcm92aWRlZCB0aGlzIGZpbGUgYXQgdGhlIHJlcXVlc3Qgb2YgYSB1c2VyIGFuZCBtYXkgaGF2ZSBjcmVhdGVkIG9yIG1vZGlmaWVkIHRoZSBmaWxlIGNvbnRlbnRzLm1zb2Z0d2FyZUFnZW50oWRuYW1lZkNsYXVkZXJhbGxBY3Rpb25zSW5jbHVkZWT1AAAAyGp1bWIAAABAanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5oYXNoLmRhdGEAAAAAGGMyc2j8DCjs08L9K/1OOlhidyYvAAAAgGNib3KlY2FsZ2ZzaGEyNTZjcGFkTQAAAAAAAAAAAAAAAABkaGFzaFggXmPTSobfFHqnVjBHmUDvvfsOCXTRuz9Dd5IHDr/KUUtkbmFtZW5qdW1iZiBtYW5pZmVzdGpleGNsdXNpb25zgaJlc3RhcnQYIWZsZW5ndGgZFooAAAI+anVtYgAAACdqdW1kYzJjbAARABCAAACqADibcQNjMnBhLmNsYWltLnYyAAAAAg9jYm9ypWNhbGdmc2hhMjU2aXNpZ25hdHVyZXhNc2VsZiNqdW1iZj0vYzJwYS91cm46YzJwYTpiZDE1MWVlNi02MzhlLTRhNDAtYmQzYS1kZThlMzgyNDc1OGIvYzJwYS5zaWduYXR1cmVqaW5zdGFuY2VJRHgseG1wOmlpZDpkOTA4N2NhZi0wOWQwLTQ4MDQtYjdjYy0zYjZkMzkyMGJmZTZyY3JlYXRlZF9hc3NlcnRpb25zg6JjdXJseC1zZWxmI2p1bWJmPWMycGEuYXNzZXJ0aW9ucy9jMnBhLmluZ3JlZGllbnQudjNkaGFzaFggykU18so7UHHNzrZAafInH/jV82RgBcP2H/mldZlNqzyiY3VybHgqc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5hY3Rpb25zLnYyZGhhc2hYIEzTgPu5Fze8N6TcsFOAS5dfuTwdtQBDM/BAKlTT/WWHomN1cmx4KXNlbGYjanVtYmY9YzJwYS5hc3NlcnRpb25zL2MycGEuaGFzaC5kYXRhZGhhc2hYICweQFTiQc9zubXh5CY4tQnd6EQP12VhgkE7epM5ebw0dGNsYWltX2dlbmVyYXRvcl9pbmZvo2RuYW1lb0FudGhyb3BpYyBGaWxlc2d2ZXJzaW9uZTEuMC4wa3NwZWNWZXJzaW9uZTIuNC4wAAAQOGp1bWIAAAAoanVtZGMyY3MAEQAQgAAAqgA4m3EDYzJwYS5zaWduYXR1cmUAAAAQCGNib3LShFkCEqIBJhghWQIKMIICBjCCAY2gAwIBAgIUQOWgCu7COdC+uIP6BkIFPWdVEwAwCgYIKoZIzj0EAwMwSTEXMBUGA1UEChMOQW50aHJvcGljLCBQQkMxLjAsBgNVBAMTJUFudGhyb3BpYyBDb250ZW50IENyZWRlbnRpYWxzIFJvb3QgQ0EwHhcNMjYwODA3MTg0MzU2WhcNMjgwODA2MTk0MzU2WjBEMRcwFQYDVQQKEw5BbnRocm9waWMsIFBCQzEpMCcGA1UEAxMgQW50aHJvcGljIENsYXVkZSBDb250ZW50IFNpZ25pbmcwWTATBgcqhkjOPQIBBggqhkjOPQMBBwNCAASYegpry1AYBRTVNL1CpTlbROnY3dey+UrsF9C3phYrATN3ZHf93Mo8RQN0KOUuOn19P4oWNFWe5n2/She9N7eTo1gwVjAOBgNVHQ8BAf8EBAMCB4AwFQYDVR0lBA4wDAYKKwYBBAGD6F4CATAMBgNVHRMBAf8EAjAAMB8GA1UdIwQYMBaAFM5R4gSBTmRbI/jjxM+aPpzB11zCMAoGCCqGSM49BAMDA2cAMGQCMDFzHRSeAXrSy1WOzkbhPZ6Km2wGTmZ/2gK18k8BQGXyqz88Rdrz6CTX9flAnYNVxgIwcF9c3fVhqmJKpi+UhasNUMko69cyX6STPfta3Q8EjyzDjzoyrol46FP6VFHhvUcJoWNwYWRZDZ4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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"


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
        background:linear-gradient(125deg,{INK} 0%,{OCHRE_800} 55%,{OCHRE_600} 100%);
        border-radius:18px; padding:2rem 2.2rem; margin-bottom:1.6rem;
        box-shadow:0 14px 34px rgba(46,45,44,0.26);
    }}
    .amr-hero::after {{
        content:""; position:absolute; top:-40%; right:-8%;
        width:360px; height:360px; border-radius:50%;
        background:radial-gradient(circle,rgba(255,255,255,0.12) 0%,rgba(255,255,255,0) 70%);
    }}
    .amr-hero h1 {{ color:{WHITE}; font-size:2.25rem; font-weight:800; margin:0;
        letter-spacing:-0.025em; line-height:1.12; }}
    .amr-hero p {{ color:#f6ddd2; font-size:1.05rem; margin:0.5rem 0 0 0; font-weight:500; }}
    .amr-eyebrow {{ display:inline-block; color:#f3c3b2; font-size:0.78rem; font-weight:700;
        letter-spacing:0.14em; text-transform:uppercase; margin-bottom:0.5rem; }}
    .amr-logo {{ position:absolute; top:1.9rem; right:2.2rem; height:48px;
        z-index:3; opacity:0.97; }}
    @media (max-width:640px) {{ .amr-logo {{ position:static; height:40px;
        margin-bottom:0.9rem; display:block; }} }}
    .sec {{ display:flex; align-items:center; gap:0.7rem; font-size:1.5rem; font-weight:800;
        color:{OCHRE_800}; letter-spacing:-0.015em; margin:0.2rem 0 1rem 0; }}
    .sec-bar {{ width:6px; height:26px; border-radius:6px;
        background:linear-gradient(180deg,{OCHRE_500},{OCHRE_700}); display:inline-block; }}
    [data-testid="stVerticalBlockBorderWrapper"] {{
        background:transparent; border:none !important; box-shadow:none !important;
        padding:0 !important; }}
    [data-testid="stPlotlyChart"] {{ background:{WHITE}; border:1px solid {LINE};
        border-radius:14px; box-shadow:0 3px 14px rgba(46,45,44,0.05);
        padding:0.6rem 0.8rem; }}
    .stButton>button {{ width:100%; border-radius:11px; height:3.2em; border:none;
        background:linear-gradient(135deg,{OCHRE_600} 0%,{OCHRE_500} 100%);
        color:{WHITE}; font-weight:700; font-size:1.02rem;
        box-shadow:0 6px 18px rgba(230,70,38,0.30); transition:all 0.16s ease; }}
    .stButton>button:hover {{ background:linear-gradient(135deg,{OCHRE_700} 0%,{OCHRE_600} 100%);
        box-shadow:0 9px 24px rgba(230,70,38,0.40); transform:translateY(-1px); color:{WHITE}; }}
    .stDownloadButton>button {{ border-radius:11px; background:{WHITE}; color:{OCHRE_700};
        border:1.6px solid {OCHRE_500}; font-weight:700; height:3em; }}
    .stDownloadButton>button:hover {{ background:{OCHRE_600}; color:{WHITE}; border-color:{OCHRE_600}; }}
    [data-testid="stSidebar"] {{ background:{WHITE}; border-right:1px solid {LINE}; }}
    [data-testid="stMetric"] {{
        background:linear-gradient(180deg,{WHITE} 0%,{OCHRE_50} 130%);
        border:1px solid {LINE}; border-left:5px solid {OCHRE_600};
        border-radius:14px; padding:1.25rem 1.45rem; box-shadow:0 3px 12px rgba(46,45,44,0.05); }}
    [data-testid="stMetricValue"] {{ color:{OCHRE_800}; font-weight:900; font-size:2.4rem; }}
    [data-testid="stMetricLabel"] {{ color:{MUTED}; font-weight:700; font-size:1rem; }}
    [data-testid="stMetricLabel"] p {{ font-size:1rem; white-space:normal;
        overflow:visible; text-overflow:clip; line-height:1.25; }}
    [data-testid="stFileUploader"] {{ background:transparent; border:none; padding:0; }}
    [data-testid="stFileUploaderDropzone"] {{
        background:{OCHRE_50}; border:1.4px dashed {OCHRE_300}; border-radius:11px;
        padding:0.5rem 0.9rem; min-height:auto; }}
    [data-testid="stDataFrame"] {{ border-radius:12px; overflow:hidden;
        border:1px solid {LINE}; box-shadow:0 3px 14px rgba(46,45,44,0.05); }}
    .legend-chip {{ display:inline-block; width:14px; height:14px; border-radius:4px;
        margin-right:8px; vertical-align:middle; }}
    hr {{ border-color:{LINE}; }}
    .kpi {{ background:{WHITE}; border:1px solid {LINE}; border-radius:16px;
        padding:1.1rem 1.25rem; box-shadow:0 3px 14px rgba(46,45,44,0.06);
        display:flex; flex-direction:column; min-height:170px; }}
    .kpi-body {{ margin-top:auto; }}
    .kpi-num {{ font-size:3rem; font-weight:900; line-height:1; color:{INK}; }}
    .kpi-lab {{ font-size:1.2rem; font-weight:800; color:{INK}; margin-top:0.35rem; }}
    .kpi-sub {{ font-size:0.95rem; font-weight:600; color:{INK}; margin-top:0.12rem; }}
    .kpi-tag {{ align-self:flex-start; font-size:0.74rem; font-weight:800;
        letter-spacing:0.06em; text-transform:uppercase; padding:3px 10px;
        border-radius:999px; background:rgba(46,45,44,0.08); color:{INK}; }}
    .kpi-total {{ background:linear-gradient(135deg,{OCHRE_700} 0%,{OCHRE_600} 100%);
        border:none; text-align:center; justify-content:center;
        box-shadow:0 8px 22px rgba(230,70,38,0.28); }}
    .kpi-total .kpi-body {{ margin-top:0; }}
    .kpi-total .kpi-num {{ color:{WHITE}; font-size:3.5rem; }}
    .kpi-total .kpi-lab {{ color:{WHITE}; }}
    </style>
""", unsafe_allow_html=True)


def section(title):
    st.markdown(f"<div class='sec'><span class='sec-bar'></span>{title}</div>", unsafe_allow_html=True)


def style_fig(fig, height, bottom=40):
    fig.update_layout(height=height, template="simple_white", plot_bgcolor="white",
        paper_bgcolor="white", font=dict(color=INK, size=15, family="Inter"),
        margin=dict(b=bottom, t=24, l=4, r=4),
        hoverlabel=dict(font=dict(size=16, family="Inter"), bgcolor="white", bordercolor=LINE))
    fig.update_xaxes(showline=True, linewidth=1.4, linecolor=LINE, tickfont=dict(size=14))
    fig.update_yaxes(showline=True, linewidth=1.4, linecolor=LINE, gridcolor=LINE, tickfont=dict(size=14))
    return fig


# ============================================================================
# 3. AST SHEET CONFIG
# ============================================================================
AST_FIRST_ABX_COL = 13
AST_CP_COL = 1
AST_DATE_COL = 0
AST_CLINIC_COL = 3
AST_SPECIES_COL = 5
AST_BREED_COL = 6
AST_SEX_COL = 7
AST_SAMPLE_COL = 8
AST_SITE_COL = 9
AST_ISOLATE_COL = 10
AST_SKIP_SHEETS = {"MALDI ID NAME LIST"}

CANON_ABBREVS = [
    "AM10", "AMC30", "CZN 30", "CL30", "CAZ30", "CVN30", "DO30", "SXT", "ENR5",
    "MAR", "GM 10", "C30", "P10", "F/M300", "N30", "OX1", "FOX30", "CC2", "E15",
    "TIC75", "TIM85", "S5", "CN 120", "RD5", "FA10", "AN30", "IPM10", "VA30",
]

# Antibiotics that drive the alert rules.
TRIGGER_ABX = ["AMC30", "ENR5", "OX1", "FOX30"]

# Core (non-antibiotic) columns, in order. (Year is tracked internally for the
# time chart but is not shown as a column: the date already carries the year.)
CORE_COLS = ["Date Received", "CP Number", "Clinic", "Organism", "Alert", "Reason"]
# Patient / sample metadata.
META_COLS = ["Species", "Breed", "Sex", "Neutered", "Sample", "Site"]
# Full record layout (used for the Excel download).
RECORD_ORDER = CORE_COLS + META_COLS + CANON_ABBREVS

# USYD clinics to include (case-insensitive match).
USYD_CLINICS = {"UVTHS", "UVTHC"}

# Per-type antibiotic breakdown shown on hover.
ALERT_BREAKDOWN = {
    "Staphylococcus (OX/FOX)": ["OX1", "FOX30"],
    "E. coli (AMC/ENR)": ["AMC30", "ENR5"],
    "Salmonella": [],
}


def meas_col(abx):
    return AST_FIRST_ABX_COL + 2 * CANON_ABBREVS.index(abx)


def interp_col(abx):
    return AST_FIRST_ABX_COL + 1 + 2 * CANON_ABBREVS.index(abx)


def normalize_cp(value):
    if value is None:
        return None
    m = re.search(r'(\d{2}-\d{4,5})', str(value))
    return m.group(1) if m else None


def normalize_clinic(value):
    """Collapse whitespace and casing so 'SASH North Ryde' == 'SASH NORTH RYDE'."""
    if value is None:
        return "NA"
    return re.sub(r'\s+', ' ', str(value)).strip().upper()


def cell_str(row, idx):
    """Safe string read of a single cell."""
    if len(row) <= idx or row[idx] in (None, ""):
        return "NA"
    return re.sub(r'\s+', ' ', str(row[idx])).strip()


def parse_sex(row):
    """Split the SEX code into sex and neutered status.
    MN = male neutered, FS = female spayed, M/F = entire, UNK = unknown."""
    s = cell_str(row, AST_SEX_COL).upper()
    if s in ("UNK", "UNKNOWN"):
        return "Unknown", "Unknown"
    if s in ("NA", "N/A", ""):
        return "NA", "NA"
    sex = "M" if s[0] == "M" else ("F" if s[0] == "F" else s)
    neutered = "Yes" if s in ("MN", "FS", "FN", "MC", "DESEXED", "NEUTERED", "SPAYED") else "No"
    return sex, neutered


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
    """Date as YYYYMMDD (no hyphens), per Kate's requested format."""
    if v is None or v == "":
        return "NA"
    if isinstance(v, datetime):
        return v.strftime("%Y%m%d")
    try:
        dt = pd.to_datetime(v, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%Y%m%d")
    except Exception:
        pass
    return str(v)


# ============================================================================
# 4. ALERT EXTRACTION
# ============================================================================
# Rules:
#   1. Escherichia coli resistant to AMC30 or ENR5
#   2. Any Staphylococcus resistant to OX1 or FOX30 (methicillin resistance)
#   3. Any Salmonella, any profile
# Scope: UVTHS and UVTHC only.
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

            clinic = normalize_clinic(row[AST_CLINIC_COL] if len(row) > AST_CLINIC_COL else None)
            if clinic not in USYD_CLINICS:
                continue

            org = str(row[AST_ISOLATE_COL]).strip()
            genus = org.split()[0].lower() if org else ""
            cp = normalize_cp(row[AST_CP_COL])
            date = fmt_date(row[AST_DATE_COL] if len(row) > AST_DATE_COL else None)
            results = {a: resolve_cell(row, a) for a in CANON_ABBREVS}
            sex, neutered = parse_sex(row)
            meta = {
                "Species": cell_str(row, AST_SPECIES_COL),
                "Breed": cell_str(row, AST_BREED_COL),
                "Sex": sex,
                "Neutered": neutered,
                "Sample": cell_str(row, AST_SAMPLE_COL),
                "Site": cell_str(row, AST_SITE_COL),
            }

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
                reason = "Any profile"

            if alert_type:
                rec = {
                    "Year": sheet_name,
                    "Date Received": date,
                    "CP Number": f"CP {cp}" if cp else "NA",
                    "Clinic": clinic,
                    "Organism": org.capitalize(),
                    "Alert": alert_type,
                    "Reason": reason,
                }
                rec.update(meta)     # species / breed / sex / sample / site
                rec.update(results)  # full susceptibility profile
                records.append(rec)

    df = pd.DataFrame(records, columns=["Year"] + RECORD_ORDER)
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
    if v == "INTR":
        return "background-color: #BFE3F7"
    return ""


def build_excel(df):
    """Full workbook: core columns plus the complete susceptibility profile."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Clinic Alerts"
    cols = RECORD_ORDER  # excludes the internal Year helper column
    hfont = Font(bold=True, color="2E2D2C", size=11)
    hborder = Border(bottom=Side(style="medium", color="E64626"))
    fills = {"R": "FFC7CE", "I": "FFEB9C", "S": "C6EFCE", "INTR": "BFE3F7"}
    for j, name in enumerate(cols, 1):
        c = ws.cell(1, j, name)
        c.font = hfont
        c.border = hborder
        c.alignment = Alignment(horizontal="left", vertical="center")
    for i, (_, row) in enumerate(df.iterrows(), start=2):
        for j, name in enumerate(cols, 1):
            v = row[name]
            cell = ws.cell(i, j, "" if (v is None or (isinstance(v, float) and pd.isna(v))) else v)
            if name in CANON_ABBREVS:
                cell.alignment = Alignment(horizontal="center")
                key = str(v).strip()
                if key in fills:
                    cell.fill = PatternFill("solid", fgColor=fills[key])
    widths = {"Date Received": 13, "CP Number": 12, "Clinic": 14,
              "Organism": 30, "Alert": 22, "Reason": 34, "Species": 11,
              "Breed": 16, "Sex": 6, "Neutered": 10, "Sample": 13, "Site": 28}
    for j, name in enumerate(cols, 1):
        letter = ws.cell(1, j).column_letter
        ws.column_dimensions[letter].width = widths.get(name, 8)
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "D2"  # keep Date / CP Number / Clinic visible while scrolling
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================================
# 5. SIDEBAR
# ============================================================================
with st.sidebar:
    st.markdown(f"<h3 style='color:{OCHRE_800};font-size:1.3rem;font-weight:800;margin-bottom:0.5rem;'>What is flagged</h3>",
                unsafe_allow_html=True)
    st.markdown(
        f"<span class='legend-chip' style='background:{STAPH_C};'></span>**Staph** &middot; OX1 / FOX30<br><br>"
        f"<span class='legend-chip' style='background:{ECOLI_C};'></span>**E. coli** &middot; AMC30 / ENR5<br><br>"
        f"<span class='legend-chip' style='background:{SALM_C};'></span>**Salmonella** &middot; any profile",
        unsafe_allow_html=True)
    st.divider()
    st.markdown(
        f"<div style='color:{INK};font-size:0.98rem;line-height:2.1;'>"
        f"<span class='legend-chip' style='background:#C6EFCE;'></span>Susceptible (S)<br>"
        f"<span class='legend-chip' style='background:#FFEB9C;'></span>Intermediate (I)<br>"
        f"<span class='legend-chip' style='background:#FFC7CE;'></span>Resistant (R)<br>"
        f"<span class='legend-chip' style='background:#BFE3F7;'></span>Intrinsic (INTR)</div>",
        unsafe_allow_html=True)

# ============================================================================
# 6. MAIN
# ============================================================================
_logo_tag = f"<img src='{LOGO_URI}' class='amr-logo' alt='University of Sydney'/>" if LOGO_URI else ""
st.markdown(
    "<div class='amr-hero'>"
    + _logo_tag +
    "<span class='amr-eyebrow'>Sydney School of Veterinary Science</span>"
    "<h1>Clinic Alert Extraction</h1>"
    "</div>",
    unsafe_allow_html=True)

with st.container(border=True):
    section("Upload the AST logging sheet")
    ast_file = st.file_uploader("AST LOGGING sheet (Excel)", type=["xlsx"])
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
        st.warning("No UVTHS or UVTHC isolates matched the alert rules in this sheet.")
    else:
        tab_alerts, tab_detail = st.tabs(["Alerts", "Sample details"])

        # ================= TAB 1: ALERTS (the actionable list) =============
        with tab_alerts:
            n_total = len(df)
            n_staph = int((df["Alert"] == "Staphylococcus (OX/FOX)").sum())
            n_ecoli = int((df["Alert"] == "E. coli (AMC/ENR)").sum())
            n_salm = int((df["Alert"] == "Salmonella").sum())
            k1, k2, k3, k4 = st.columns(4)
            k1.markdown(
                f"<div class='kpi kpi-total'><div class='kpi-body'>"
                f"<div class='kpi-num'>{n_total}</div>"
                f"<div class='kpi-lab'>Total alerts</div></div></div>", unsafe_allow_html=True)
            k2.markdown(
                f"<div class='kpi' style='border-left:5px solid {STAPH_C}'>"
                f"<span class='kpi-tag'>Resistant</span><div class='kpi-body'>"
                f"<div class='kpi-num'>{n_staph}</div>"
                f"<div class='kpi-lab'>Staph</div><div class='kpi-sub'>OX1 / FOX30</div>"
                f"</div></div>", unsafe_allow_html=True)
            k3.markdown(
                f"<div class='kpi' style='border-left:5px solid {ECOLI_C}'>"
                f"<span class='kpi-tag'>Resistant</span><div class='kpi-body'>"
                f"<div class='kpi-num'>{n_ecoli}</div>"
                f"<div class='kpi-lab'>E. coli</div><div class='kpi-sub'>AMC30 / ENR5</div>"
                f"</div></div>", unsafe_allow_html=True)
            k4.markdown(
                f"<div class='kpi' style='border-left:5px solid {SALM_C}'>"
                f"<span class='kpi-tag'>All isolates</span><div class='kpi-body'>"
                f"<div class='kpi-num'>{n_salm}</div>"
                f"<div class='kpi-lab'>Salmonella</div><div class='kpi-sub'>any profile</div>"
                f"</div></div>", unsafe_allow_html=True)
            st.write("")

            with st.container(border=True):
                section("Flagged isolates")
                f1, f2 = st.columns(2)
                with f1:
                    pick_clinic = st.selectbox("Clinic", ["All clinics"] + sorted(df["Clinic"].unique()), key="a_clinic")
                with f2:
                    pick_type = st.selectbox("Alert type", ["All types"] + sorted(df["Alert"].unique()), key="a_type")
                view = df.copy()
                if pick_clinic != "All clinics":
                    view = view[view["Clinic"] == pick_clinic]
                if pick_type != "All types":
                    view = view[view["Alert"] == pick_type]
                view = view.reset_index(drop=True)
                view.index = range(1, len(view) + 1)

                st.markdown(f"**{len(view)} isolate(s) shown**")
                st.dataframe(view[CORE_COLS + TRIGGER_ABX].style.map(sir_css, subset=TRIGGER_ABX),
                             use_container_width=True)

                fname = f"Clinic_Alerts_{datetime.utcnow().strftime('%Y%m%d')}.xlsx"
                st.download_button("Download full record (Excel)", build_excel(view), fname,
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # ============= TAB 2: SAMPLE DETAILS + ANALYTICS ===================
        with tab_detail:
            with st.container(border=True):
                section("Alerts by clinic")
                cc = df["Clinic"].value_counts().sort_values(ascending=True)
                ys, xs = cc.index.tolist(), [int(v) for v in cc.values]
                figc = go.Figure(go.Bar(orientation="h", y=ys, x=xs, marker_color=OCHRE_600,
                    marker_line_width=0, text=xs, textposition="outside",
                    textfont=dict(size=14, color=INK), cliponaxis=False,
                    hovertemplate="<b>%{y}</b><br>Alerts: %{x}<extra></extra>"))
                style_fig(figc, max(200, 60 * len(ys)), bottom=30)
                figc.update_layout(xaxis_title="<b>Number of alerts</b>", yaxis_title="")
                figc.update_xaxes(range=[0, (max(xs) if xs else 1) * 1.2])
                st.plotly_chart(figc, use_container_width=True, key="chart_clinic")

            st.write("")

            with st.container(border=True):
                section("Alerts over time")
                years = sorted(df["Year"].unique())
                figt = go.Figure()
                for atype, color in ALERT_COLORS.items():
                    counts = [int(((df["Year"] == y) & (df["Alert"] == atype)).sum()) for y in years]
                    breakdown = ALERT_BREAKDOWN[atype]
                    ht = "<b>%{x}</b><br>" + atype + ": %{y}"
                    if breakdown:
                        customdata = []
                        for y in years:
                            sub = df[(df["Year"] == y) & (df["Alert"] == atype)]
                            customdata.append([int((sub[a] == "R").sum()) for a in breakdown])
                        for i, a in enumerate(breakdown):
                            ht += f"<br>{a} resistant: %{{customdata[{i}]}}"
                        figt.add_bar(x=years, y=counts, name=atype, marker_color=color,
                            customdata=customdata, hovertemplate=ht + "<extra></extra>")
                    else:
                        figt.add_bar(x=years, y=counts, name=atype, marker_color=color,
                            hovertemplate=ht + "<extra></extra>")
                style_fig(figt, 430, bottom=40)
                figt.update_layout(barmode="group", xaxis_title="<b>Year</b>",
                    yaxis_title="<b>Number of alerts</b>",
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=13)))
                figt.update_xaxes(type="category")
                st.plotly_chart(figt, use_container_width=True, key="chart_time")

            st.write("")

            with st.container(border=True):
                section("Alerts by species")
                sp = df["Species"].value_counts().sort_values(ascending=True)
                ys, xs = sp.index.tolist(), [int(v) for v in sp.values]
                figs = go.Figure(go.Bar(orientation="h", y=ys, x=xs, marker_color=SALM_C,
                    marker_line_width=0, text=xs, textposition="outside",
                    textfont=dict(size=14, color=INK), cliponaxis=False,
                    hovertemplate="<b>%{y}</b><br>Alerts: %{x}<extra></extra>"))
                style_fig(figs, max(200, 52 * len(ys)), bottom=30)
                figs.update_layout(xaxis_title="<b>Number of alerts</b>", yaxis_title="")
                figs.update_xaxes(range=[0, (max(xs) if xs else 1) * 1.2])
                st.plotly_chart(figs, use_container_width=True, key="chart_species")

            st.write("")

            with st.container(border=True):
                section("Sample details")
                g1, g2 = st.columns(2)
                with g1:
                    pick_clinic2 = st.selectbox("Clinic", ["All clinics"] + sorted(df["Clinic"].unique()), key="d_clinic")
                with g2:
                    pick_type2 = st.selectbox("Alert type", ["All types"] + sorted(df["Alert"].unique()), key="d_type")
                dview = df.copy()
                if pick_clinic2 != "All clinics":
                    dview = dview[dview["Clinic"] == pick_clinic2]
                if pick_type2 != "All types":
                    dview = dview[dview["Alert"] == pick_type2]
                dview = dview.reset_index(drop=True)
                dview.index = range(1, len(dview) + 1)
                detail_cols = ["CP Number", "Organism", "Alert"] + META_COLS
                st.markdown(f"**{len(dview)} isolate(s) shown**")
                st.dataframe(dview[detail_cols], use_container_width=True)
else:
    st.info("Upload the AST logging sheet to begin.")
