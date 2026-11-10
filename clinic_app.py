import streamlit as st
import pandas as pd
import re
import io
import os
import base64
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
OCHRE_50 = "#fdf1ec"; INK = "#141414"; MUTED = "#141414"
LINE = "#e3ded9"; CANVAS = "#ffffff"; WHITE = "#ffffff"

# One ochre family, dimmed (muted dark / mid / light), not three bright hues.
STAPH_C = "#8c3b2b"   # dimmed dark ochre
ECOLI_C = "#c47050"   # dimmed mid ochre
SALM_C = "#e6bda9"    # dimmed light ochre

ALERT_COLORS = {
    "Staphylococcus (OX/FOX)": STAPH_C,
    "E. coli (AMC/ENR)": ECOLI_C,
    "Salmonella": SALM_C,
}

# Full organism names for legends and hovers.
ALERT_FULL = {
    "Staphylococcus (OX/FOX)": "Staphylococcus",
    "E. coli (AMC/ENR)": "E. coli",
    "Salmonella": "Salmonella",
}
ALERT_ORDER = ["Staphylococcus (OX/FOX)", "E. coli (AMC/ENR)", "Salmonella"]

def _load_logo(name="usyd_logo_white.png"):
    """Inline the USYD logo committed alongside this app as a data URI."""
    for candidate in (name, os.path.join(os.path.dirname(os.path.abspath(__file__)), name)):
        try:
            with open(candidate, "rb") as f:
                return "data:image/png;base64," + base64.b64encode(f.read()).decode()
        except Exception:
            continue
    return ""


LOGO_URI = _load_logo()


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
        background:{OCHRE_600};
        border-radius:18px; padding:2rem 2.2rem; margin-bottom:1.6rem;
        box-shadow:0 8px 22px rgba(230,70,38,0.20);
    }}
    .amr-hero h1 {{ color:{WHITE}; font-size:2.35rem; font-weight:800; margin:0;
        letter-spacing:-0.025em; line-height:1.12; }}
    .amr-eyebrow {{ display:inline-block; color:#ffe3da; font-size:0.82rem; font-weight:700;
        letter-spacing:0.14em; text-transform:uppercase; margin-bottom:0.5rem; }}
    .amr-logo {{ position:absolute; top:1.9rem; right:2.2rem; height:48px;
        z-index:3; opacity:0.97; }}
    @media (max-width:640px) {{ .amr-logo {{ position:static; height:40px;
        margin-bottom:0.9rem; display:block; }} }}
    .sec {{ display:flex; align-items:center; gap:0.7rem; font-size:1.5rem; font-weight:800;
        color:{OCHRE_800}; letter-spacing:-0.015em; margin:0.2rem 0 1rem 0; }}
    .sec-bar {{ width:6px; height:26px; border-radius:6px;
        background:{OCHRE_600}; display:inline-block; }}
    [data-testid="stVerticalBlockBorderWrapper"] {{
        background:transparent; border:none !important; box-shadow:none !important;
        padding:0 !important; }}
    [data-testid="stPlotlyChart"] {{ background:{WHITE}; border:1px solid {LINE};
        border-radius:14px; box-shadow:0 3px 14px rgba(46,45,44,0.05);
        padding:0.6rem 0.8rem; }}
    .stButton>button {{ width:100%; border-radius:11px; height:3.2em; border:none;
        background:{OCHRE_600};
        color:{WHITE}; font-weight:700; font-size:1.05rem;
        box-shadow:0 4px 12px rgba(230,70,38,0.22); transition:all 0.16s ease; }}
    .stButton>button:hover {{ background:{OCHRE_700};
        box-shadow:0 6px 16px rgba(230,70,38,0.30); transform:translateY(-1px); color:{WHITE}; }}
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
    .sb-row {{ margin-bottom:0.85rem; }}
    .sb-tag {{ font-size:0.66rem; font-weight:800; text-transform:uppercase;
        letter-spacing:0.05em; padding:1px 7px; border-radius:999px;
        background:{OCHRE_50}; color:{OCHRE_800}; }}
    .sb-sub {{ display:block; font-size:0.86rem; color:#8a7f79;
        margin-left:22px; margin-top:1px; }}
    .sb-key {{ color:{OCHRE_800}; font-size:1.05rem; font-weight:800; margin:0.2rem 0 0.5rem 0; }}
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
        border-radius:999px; background:{OCHRE_50}; color:{OCHRE_800}; }}
    .kpi-total {{ background:{OCHRE_600};
        border:none; text-align:center; justify-content:center;
        box-shadow:0 6px 16px rgba(230,70,38,0.24); }}
    .kpi-total .kpi-body {{ margin-top:0; }}
    .kpi-total .kpi-num {{ color:{WHITE}; font-size:3.5rem; }}
    .kpi-total .kpi-lab {{ color:{WHITE}; }}
    </style>
""", unsafe_allow_html=True)


def section(title):
    st.markdown(f"<div class='sec'><span class='sec-bar'></span>{title}</div>", unsafe_allow_html=True)


def style_fig(fig, height, bottom=40):
    fig.update_layout(height=height, template="simple_white", plot_bgcolor="white",
        paper_bgcolor="white", font=dict(color=INK, size=17, family="Inter"),
        margin=dict(b=bottom, t=24, l=4, r=4),
        hoverlabel=dict(font=dict(size=16, color=INK, family="Inter"), bgcolor="white", bordercolor=LINE))
    fig.update_xaxes(showline=True, linewidth=1.4, linecolor="#9a948e",
        tickfont=dict(size=16, color=INK), title_font=dict(size=16, color=INK))
    fig.update_yaxes(showline=True, linewidth=1.4, linecolor="#9a948e", gridcolor=LINE,
        tickfont=dict(size=16, color=INK), title_font=dict(size=16, color=INK))
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

# Disk code -> full antibiotic name for display (clinicians don't read disk codes).
ABX_FULL = {
    "AM10": "Ampicillin", "AMC30": "Amoxicillin-clavulanate", "CZN 30": "Cefazolin",
    "CL30": "Cefalexin", "CAZ30": "Ceftazidime", "CVN30": "Cefovecin", "DO30": "Doxycycline",
    "SXT": "Trimethoprim-sulfamethoxazole", "ENR5": "Enrofloxacin", "MAR": "Marbofloxacin",
    "GM 10": "Gentamicin", "C30": "Chloramphenicol", "P10": "Penicillin",
    "F/M300": "Nitrofurantoin", "N30": "Neomycin", "OX1": "Oxacillin", "FOX30": "Cefoxitin",
    "CC2": "Clindamycin", "E15": "Erythromycin", "TIC75": "Ticarcillin",
    "TIM85": "Ticarcillin-clavulanate", "S5": "Streptomycin", "CN 120": "Gentamicin (high-level)",
    "RD5": "Rifampicin", "FA10": "Fusidic acid", "AN30": "Amikacin", "IPM10": "Imipenem",
    "VA30": "Vancomycin",
}

# Antibiotics that drive the alert rules.
TRIGGER_ABX = ["AMC30", "ENR5", "OX1", "FOX30"]

# Core (non-antibiotic) columns, in order. (Year is tracked internally for the
# time chart but is not shown as a column: the date already carries the year.)
CORE_COLS = ["Date Received", "CP Number", "Clinic", "Organism", "Alert", "Reason"]
# Patient / sample metadata.
META_COLS = ["Species", "Breed", "Sex", "Neutered", "Sample", "Site", "Sample Site (Detailed)"]
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


# Different wordings for the same anatomical site, collapsed onto one label.
# Keys are lowercase; add or adjust entries here as Kate confirms them.
SITE_SYNONYMS = {
    "foot": "Foot", "feet": "Foot",
    "belly": "Abdomen", "ventrum": "Abdomen",
    "back": "Dorsum", "dorsum": "Dorsum",
    "axilla": "Axilla", "axila": "Axilla",
    "nail bed": "Nail bed", "nailbed": "Nail bed",
}

# Anatomical direction / plane words. When one of these sits in front of a site word
# ("Ventral abdomen", "Lateral carpus") it is a modifier: the word after it is the
# site and the direction moves to the detailed column. Laterality is handled here too.
SITE_MODIFIERS = {
    "lt": "Left", "rt": "Right",
    "left": "Left", "right": "Right", "bilateral": "Bilateral",
    "ventral": "Ventral", "dorsal": "Dorsal", "lateral": "Lateral",
    "medial": "Medial", "cranial": "Cranial", "caudal": "Caudal",
    "rostral": "Rostral", "proximal": "Proximal", "distal": "Distal",
}
LATERALITY = {"Left", "Right", "Bilateral"}
_MOD_RE = re.compile(r'\b(' + '|'.join(SITE_MODIFIERS) + r')\b', flags=re.IGNORECASE)

# Lesion / finding words (not an anatomical site) and procedure / sample-handling
# words. These never win over a real site: when the main text is only these and a
# real site sits alongside, the site is kept and these move to the detailed column.
LESION_WORDS = {
    "abscess", "sinus", "discharging", "draining", "discharge", "mass", "lesion",
    "wound", "ulcer", "ulcerated", "nodule", "swelling", "cyst", "fistula", "pustule",
    "pus", "exudate", "lump", "growth", "tract", "tracts", "scab", "crust", "erosion",
    "papule", "purulent",
}
PROCEDURE_WORDS = {
    "flush", "flushed", "flushing", "debridement", "debride", "debrided", "lavage",
    "biopsy", "aspirate", "aspiration", "wash", "washed", "voided", "cysto", "op",
    "surgical", "repair", "culture", "exit", "pre", "post", "free", "catch", "punch",
    "swab", "swabs", "gel", "plain", "tube", "site", "area", "screw", "screws",
    "from", "and", "teca", "tecalo", "tecabo", "bo",
}
_QUALIFIER = LESION_WORDS | PROCEDURE_WORDS
# Separators split site from detail: comma / colon / semicolon / pipe / slash, or a
# SPACED hyphen. A word-internal hyphen (PRE-LAVAGE, PERI-ORAL) is NOT a split point.
_SPLIT_RE = re.compile(r'\s*[,:;|/]\s*|\s+-\s+')
_TOKEN_RE = re.compile(r"[A-Za-z]+")
_COUNT_RE = re.compile(r'^\d+\s*x$', flags=re.IGNORECASE)


def _strippable(word):
    """A count prefix (3X) or a word made only of lesion / procedure tokens."""
    if _COUNT_RE.match(word):
        return True
    alts = [a.lower() for a in _TOKEN_RE.findall(word)]
    return bool(alts) and all(a in _QUALIFIER for a in alts)


def _classify(text):
    core = _MOD_RE.sub(' ', text)
    toks = [t.lower() for t in _TOKEN_RE.findall(core)]
    if not toks:
        return 'modifier' if _TOKEN_RE.findall(text) else 'empty'
    if all(t in LESION_WORDS for t in toks):
        return 'lesion'
    if all(t in PROCEDURE_WORDS for t in toks):
        return 'procedure'
    return 'site'


def _cap_detail(text):
    text = text.strip(" ,;-/&+")
    if not text:
        return "NA"
    return ", ".join(p.strip().capitalize() for p in text.split(",") if p.strip()) or "NA"


def _cap_phrase(text):
    """Whole phrase in sentence case, kept as one unit (not comma-split)."""
    text = re.sub(r'\s+', ' ', re.sub(r'[&/+]', ' ', text)).strip(" ,;-/&+")
    return (text[0].upper() + text[1:].lower()) if text else "NA"


def _join(a, b):
    parts = [p for p in (a, b) if p and p != "NA"]
    return ", ".join(parts)


def clean_site(raw):
    """Main site + detail from a free-text site field. A parenthetical or post-separator
    fragment becomes the detail; the part that is a real anatomical site is kept while
    lesion / procedure / laterality words move to the detailed column."""
    s = str(raw).strip() if raw not in (None, "") else ""
    if not s or s.upper() in ("NA", "N/A"):
        return "NA", "NA"
    # A whole field that is only a collection method / procedure (CYSTO, FREE CATCH,
    # PUNCH BIOPSY, VOIDED) is not a site: keep it, as one phrase, in the detail.
    if "(" not in s and not _SPLIT_RE.search(s) and _classify(s) == "procedure":
        return "NA", _cap_phrase(s)
    detail = ""
    m = re.search(r'^(.*?)\s*\((.*?)\)', s)
    if m and (m.group(1).strip() or m.group(2).strip()):
        main = m.group(1).strip()
        after = s[m.end():].strip()
        detail = m.group(2).strip()
        s = main or after
    else:
        parts = _SPLIT_RE.split(s, maxsplit=1)
        if len(parts) > 1:
            s = parts[0].strip()
            detail = re.sub(r'[\)\]\}]+$', '', parts[1].strip()).strip()

    # Of the two parts, keep whichever is the better anatomical site.
    if detail:
        rank = {'site': 3, 'lesion': 2, 'procedure': 1, 'modifier': 0, 'empty': 0}
        if rank[_classify(detail)] > rank[_classify(s)]:
            s, detail = detail, s

    # Strip leading / trailing qualifier words (POST-OP ABDOMINAL CAVITY -> Abdominal).
    words = s.split()
    lead = []
    while len(words) > 1 and _strippable(words[0]):
        lead.append(words.pop(0))
    trail = []
    while len(words) > 1 and _strippable(words[-1]):
        trail.insert(0, words.pop())
    if lead or trail:
        s = " ".join(words)
        detail = _join(detail, " ".join(lead + trail))

    # Direction / laterality words are modifiers: the word after them is the site.
    found = [SITE_MODIFIERS[w.lower()] for w in _MOD_RE.findall(s)]
    core = re.sub(r'[&/+]', ' ', _MOD_RE.sub(' ', s))
    core = re.sub(r'\s+', ' ', core).strip(" -,&/+")
    mod_text = ", ".join(dict.fromkeys(found))
    if found:
        if core:
            detail = _join(mod_text, detail)
            s = core
        elif all(f in LATERALITY for f in found):
            detail = _join(mod_text, detail)   # lone laterality is not a site
            s = ""
        # else: keep a bare plane word (e.g. "Ventral") as the site

    s = re.sub(r'^[^0-9A-Za-z]+|[^0-9A-Za-z]+$', '', s).strip()
    cls = _classify(s) if s else 'empty'
    if cls in ('procedure', 'empty') or (cls == 'modifier' and s.capitalize() in LATERALITY):
        if s and cls != 'empty':
            detail = _join(detail, s)
        s = ""

    s = (s[0].upper() + s[1:].lower()) if s else "NA"
    s = SITE_SYNONYMS.get(s.lower().strip(), s)
    detail = _cap_detail(detail)
    return s, detail


BREED_SYNONYMS = {
    "labrador": "Labrador Retriever",
    "english cocker spaniel": "Cocker Spaniel",
    "dsh": "Domestic Shorthair",
    "dlh": "Domestic Longhair",
    "domestic short hair": "Domestic Shorthair",
    "domestic long hair": "Domestic Longhair",
    "staffy": "Staffordshire Bull Terrier",
    "american staff terrier": "American Staffordshire Terrier",
    "american staffy": "American Staffordshire Terrier",
    "amstaff": "American Staffordshire Terrier",
    "neopolitan mastiff": "Neapolitan Mastiff",
}


def normalize_breed(breed):
    """Tidy a breed so variants match: title-case (capital first letter, rest lower),
    standardise the crossbreed marker by dropping a trailing ' X', and map known
    synonyms and abbreviations (DSH, Labrador, American Staff terrier...) to one name."""
    if breed is None:
        return "NA"
    b = re.sub(r'\s+', ' ', str(breed)).strip()
    if not b or b.upper() in ("NA", "N/A"):
        return "NA"
    base = re.sub(r'[\s(]*\b(x|cross|crossbreed)\b[\s)]*$', '', b, flags=re.IGNORECASE).strip()
    return BREED_SYNONYMS.get(base.lower(), base.title())


def normalize_species(species):
    """Species in Title Case so 'CANINE' and 'canine' both read 'Canine'."""
    if species is None:
        return "NA"
    s = re.sub(r'\s+', ' ', str(species)).strip()
    if not s or s.upper() in ("NA", "N/A"):
        return "NA"
    return s.title()


# Bacterial isolate name normalization: fixes spelling variants (and casing) so the
# same organism does not appear as several isolate types. Genus misspellings also
# matter for the alert rules, which match on the genus. Edit these maps as needed.
GENUS_FIX = {
    "staphyloccus": "Staphylococcus", "staphylococus": "Staphylococcus",
    "staphyloccocus": "Staphylococcus", "staphylococccus": "Staphylococcus",
    "streptococccus": "Streptococcus", "streptoccocus": "Streptococcus",
    "streptoccoccus": "Streptococcus", "streptococus": "Streptococcus",
    "pseudomona": "Pseudomonas", "pseudomonos": "Pseudomonas",
    "escherchia": "Escherichia", "echerichia": "Escherichia",
    "enterococccus": "Enterococcus", "enterococus": "Enterococcus",
}
SPECIES_FIX = {"dysagalactiae": "dysgalactiae", "pseudointermedius": "pseudintermedius"}
# Old name -> current name, by genus (S. intermedius group in companion animals).
SPECIES_MERGE = {("staphylococcus", "intermedius"): "pseudintermedius"}


def normalize_organism(org):
    if org is None:
        return "NA"
    o = re.sub(r'\s+', ' ', str(org)).strip()
    if not o or o.upper() in ("NA", "N/A"):
        return "NA"
    words = o.split()
    genus = GENUS_FIX.get(words[0].lower(), words[0].capitalize())
    gkey = genus.lower()
    out = [genus]
    for i, w in enumerate(words[1:]):
        parts = []
        for p in w.split("/"):          # species may be slash-joined alternatives
            pl = SPECIES_FIX.get(p.lower(), p.lower())
            if i == 0:
                pl = SPECIES_MERGE.get((gkey, pl), pl)
            parts.append(pl)
        out.append("/".join(parts))
    return " ".join(out)


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

            org = normalize_organism(str(row[AST_ISOLATE_COL]).strip())
            genus = org.split()[0].lower() if org and org != "NA" else ""
            cp = normalize_cp(row[AST_CP_COL])
            date = fmt_date(row[AST_DATE_COL] if len(row) > AST_DATE_COL else None)
            results = {a: resolve_cell(row, a) for a in CANON_ABBREVS}
            sex, neutered = parse_sex(row)
            site_main, site_detail = clean_site(cell_str(row, AST_SITE_COL))
            meta = {
                "Species": normalize_species(cell_str(row, AST_SPECIES_COL)),
                "Breed": normalize_breed(cell_str(row, AST_BREED_COL)),
                "Sex": sex,
                "Neutered": neutered,
                "Sample": cell_str(row, AST_SAMPLE_COL),
                "Site": site_main,
                "Sample Site (Detailed)": site_detail,
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
                    "Organism": org,
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
        is_abx = name in CANON_ABBREVS
        c = ws.cell(1, j, ABX_FULL.get(name, name) if is_abx else name)
        c.font = hfont
        c.border = hborder
        if is_abx:   # full drug name, rotated so the sheet stays compact
            c.alignment = Alignment(text_rotation=90, horizontal="center", vertical="bottom")
        else:
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
              "Breed": 16, "Sex": 6, "Neutered": 10, "Sample": 13, "Site": 22,
              "Sample Site (Detailed)": 24}
    for j, name in enumerate(cols, 1):
        letter = ws.cell(1, j).column_letter
        ws.column_dimensions[letter].width = 5.2 if name in CANON_ABBREVS else widths.get(name, 8)
    abx_hdr_len = max([len(ABX_FULL.get(a, a)) for a in CANON_ABBREVS if a in cols] + [0])
    ws.row_dimensions[1].height = min(max(abx_hdr_len * 6.2, 22), 215)
    ws.freeze_panes = "D2"  # keep Date / CP Number / Clinic visible while scrolling
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================================
# 5. SIDEBAR
# ============================================================================
with st.sidebar:
    st.markdown(f"<h3 style='color:{OCHRE_800};font-size:1.3rem;font-weight:800;margin-bottom:0.7rem;'>Alerts</h3>",
                unsafe_allow_html=True)
    st.markdown(
        f"<div class='sb-row'><span class='legend-chip' style='background:{STAPH_C};'></span>"
        f"<b>Staphylococcus</b> <span class='sb-tag'>Resistant</span>"
        f"<span class='sb-sub'>OX1 / FOX30</span></div>"
        f"<div class='sb-row'><span class='legend-chip' style='background:{ECOLI_C};'></span>"
        f"<b>E. coli</b> <span class='sb-tag'>Resistant</span>"
        f"<span class='sb-sub'>AMC30 / ENR5</span></div>"
        f"<div class='sb-row'><span class='legend-chip' style='background:{SALM_C};'></span>"
        f"<b>Salmonella</b> <span class='sb-tag'>All</span>"
        f"<span class='sb-sub'>Any profile</span></div>",
        unsafe_allow_html=True)
    st.divider()
    st.markdown("<div class='sb-key'>Colour key</div>", unsafe_allow_html=True)
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
    section("Upload the AST logging sheet (Excel)")
    ast_file = st.file_uploader("AST logging sheet", type=["xlsx"], label_visibility="collapsed")
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
        tab_alerts, tab_analytics = st.tabs(["Alerts", "Analytics"])

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
                f"<div class='kpi-lab'>Staphylococcus</div><div class='kpi-sub'>OX1 / FOX30</div>"
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
                f"<div class='kpi-lab'>Salmonella</div><div class='kpi-sub'>Any profile</div>"
                f"</div></div>", unsafe_allow_html=True)
            st.write("")

            with st.container(border=True):
                section("Flagged isolates")
                f1, f2, f3 = st.columns(3)
                with f1:
                    pick_clinic = st.selectbox("Clinic", ["All clinics"] + sorted(df["Clinic"].unique()), key="a_clinic")
                with f2:
                    pick_type = st.selectbox("Alert type", ["All types"] + sorted(df["Alert"].unique()), key="a_type")
                with f3:
                    pick_year = st.selectbox("Year", ["All years"] + sorted(df["Year"].unique()), key="a_year")
                view = df.copy()
                if pick_clinic != "All clinics":
                    view = view[view["Clinic"] == pick_clinic]
                if pick_type != "All types":
                    view = view[view["Alert"] == pick_type]
                if pick_year != "All years":
                    view = view[view["Year"] == pick_year]
                view = view.reset_index(drop=True)
                view.index = range(1, len(view) + 1)

                st.markdown(f"**{len(view)} isolate(s) shown**")
                show_cols = CORE_COLS + META_COLS + CANON_ABBREVS
                shown = view[show_cols].rename(columns={c: ABX_FULL.get(c, c) for c in CANON_ABBREVS})
                subset_full = [ABX_FULL.get(c, c) for c in CANON_ABBREVS]
                st.dataframe(shown.style.map(sir_css, subset=subset_full), use_container_width=True)

                fname = f"Clinic_Alerts_{datetime.utcnow().strftime('%Y%m%d')}.xlsx"
                st.download_button("Download full record (Excel)", build_excel(view), fname,
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # ============= TAB 2: ANALYTICS (charts only) ======================
        def _breakdown_hover(groups):
            """Hover text + customdata: total next to the group name, then a
            per-alert-type count underneath, using full organism names."""
            cust = []
            for g in groups:
                cust.append([int(v) for v in g])
            ht = "<b>%{y}: %{x}</b>"
            for i, t in enumerate(ALERT_ORDER):
                ht += f"<br>{ALERT_FULL[t]}: %{{customdata[{i}]}}"
            return ht + "<extra></extra>", cust

        with tab_analytics:
            # ---- Alerts by clinic (hover: breakdown by organism) ----
            with st.container(border=True):
                section("Alerts by clinic")
                cc = df["Clinic"].value_counts().sort_values(ascending=True)
                ys, xs = cc.index.tolist(), [int(v) for v in cc.values]
                groups = [[int((df[(df["Clinic"] == cl)]["Alert"] == t).sum()) for t in ALERT_ORDER] for cl in ys]
                ht, cust = _breakdown_hover(groups)
                figc = go.Figure(go.Bar(orientation="h", y=ys, x=xs, marker_color=OCHRE_600,
                    marker_line_width=0, text=xs, textposition="outside",
                    textfont=dict(size=16, color=INK), cliponaxis=False,
                    customdata=cust, hovertemplate=ht))
                style_fig(figc, max(200, 66 * len(ys)), bottom=34)
                figc.update_layout(xaxis_title="<b>Number of alerts</b>", yaxis_title="")
                figc.update_xaxes(range=[0, (max(xs) if xs else 1) * 1.2],
                                  title_font=dict(size=16, color=INK))
                st.plotly_chart(figc, use_container_width=True, key="chart_clinic")

            st.divider()

            # ---- Alerts over time (grouped by organism, hover shows R counts) ----
            with st.container(border=True):
                section("Alerts over time")
                years = sorted(df["Year"].unique())
                figt = go.Figure()
                for atype in ALERT_ORDER:
                    color = ALERT_COLORS[atype]
                    name = ALERT_FULL[atype]
                    counts = [int(((df["Year"] == y) & (df["Alert"] == atype)).sum()) for y in years]
                    breakdown = ALERT_BREAKDOWN[atype]
                    ht = "<b>" + name + ": %{y}</b>"
                    if breakdown:
                        # Non-overlapping split so the parts add up to the total.
                        a, b = breakdown
                        customdata = []
                        for y in years:
                            sub = df[(df["Year"] == y) & (df["Alert"] == atype)]
                            ra, rb = (sub[a] == "R"), (sub[b] == "R")
                            customdata.append([int((ra & rb).sum()), int((ra & ~rb).sum()), int((~ra & rb).sum())])
                        ht += (f"<br>{a} + {b}: %{{customdata[0]}}"
                               f"<br>{a} only: %{{customdata[1]}}"
                               f"<br>{b} only: %{{customdata[2]}}")
                        figt.add_bar(x=years, y=counts, name=name, marker_color=color,
                            customdata=customdata, hovertemplate=ht + "<extra></extra>")
                    else:
                        figt.add_bar(x=years, y=counts, name=name, marker_color=color,
                            hovertemplate=ht + "<extra></extra>")
                style_fig(figt, 440, bottom=44)
                figt.update_layout(barmode="group", xaxis_title="<b>Year</b>",
                    yaxis_title="<b>Number of alerts</b>",
                    legend=dict(orientation="v", yanchor="top", y=0.99, xanchor="right", x=0.995,
                        font=dict(size=14, color=INK), bgcolor="rgba(255,255,255,0.88)",
                        bordercolor=LINE, borderwidth=1))
                figt.update_xaxes(type="category", title_font=dict(size=16, color=INK))
                figt.update_yaxes(title_font=dict(size=16, color=INK))
                st.plotly_chart(figt, use_container_width=True, key="chart_time")

            st.divider()

            # ---- Alerts by species (hover: breakdown by organism) ----
            with st.container(border=True):
                section("Alerts by species")
                sp = df["Species"].value_counts().sort_values(ascending=True)
                ys, xs = sp.index.tolist(), [int(v) for v in sp.values]
                groups = [[int((df[(df["Species"] == s)]["Alert"] == t).sum()) for t in ALERT_ORDER] for s in ys]
                ht, cust = _breakdown_hover(groups)
                figs = go.Figure(go.Bar(orientation="h", y=ys, x=xs, marker_color=OCHRE_600,
                    marker_line_width=0, text=xs, textposition="outside",
                    textfont=dict(size=16, color=INK), cliponaxis=False,
                    customdata=cust, hovertemplate=ht))
                style_fig(figs, max(200, 58 * len(ys)), bottom=34)
                figs.update_layout(xaxis_title="<b>Number of alerts</b>", yaxis_title="")
                figs.update_xaxes(range=[0, (max(xs) if xs else 1) * 1.2],
                                  title_font=dict(size=16, color=INK))
                st.plotly_chart(figs, use_container_width=True, key="chart_species")

            st.divider()

            # ---- Breed prevalence by host species (one bar chart per species) ----
            def _org_breakdown(frame, top=6):
                """Organisms flagged in a chart slice, as a hover list with counts."""
                vc = frame["Organism"].value_counts()
                items = [f"{name} ({int(c)})" for name, c in vc.items()][:top]
                if len(vc) > top:
                    items.append(f"+{len(vc) - top} more")
                return "<br>".join(items) if items else "—"

            with st.container(border=True):
                section("Breed prevalence by host species")
                st.caption("One count per flagged animal.")
                unique_demo = df.drop_duplicates(subset=["CP Number"])
                host_species = [s for s in unique_demo["Species"].unique() if s not in ("NA", "nan")]
                if host_species:
                    for host in host_species:
                        sub = unique_demo[unique_demo["Species"] == host]
                        b_counts = (sub[sub["Breed"] != "NA"]["Breed"]
                                    .value_counts().sort_values(ascending=True))
                        if b_counts.empty:
                            continue
                        yb = b_counts.index.tolist()
                        xb = [int(v) for v in b_counts.values]
                        b_cust = [[_org_breakdown(df[(df["Species"] == host) &
                                                     (df["Breed"] == breed)])] for breed in yb]
                        st.markdown(
                            f"<div style='color:{OCHRE_800};font-size:1.05rem;font-weight:800;"
                            f"margin:0.5rem 0 0.15rem;'>{host} breeds "
                            f"<span style='font-weight:600;'>({len(sub)} animals)</span></div>",
                            unsafe_allow_html=True)
                        figb = go.Figure(go.Bar(orientation="h", y=yb, x=xb, marker_color=OCHRE_800,
                            marker_line_width=0, text=xb, textposition="outside",
                            textfont=dict(size=16, color=INK), cliponaxis=False, customdata=b_cust,
                            hovertemplate="<b>%{y}</b><br>Animals: %{x}"
                                          "<br><br><b>Organisms flagged</b><br>%{customdata[0]}<extra></extra>"))
                        style_fig(figb, max(160, 34 * len(yb)), bottom=30)
                        figb.update_layout(xaxis_title="<b>Number of animals</b>", yaxis_title="")
                        figb.update_xaxes(range=[0, (max(xb) if xb else 1) * 1.2], dtick=1,
                                          title_font=dict(size=16, color=INK))
                        st.plotly_chart(figb, use_container_width=True, key=f"chart_breed_{host}")
                else:
                    st.info("No host-species demographic data available.")
else:
    st.info("Upload the AST logging sheet to begin.")
