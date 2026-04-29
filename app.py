"""
ZÜS PDF Rename Pipeline — Streamlit Web App
============================================
Upload ZÜS inspection PDFs → get automatically renamed files + XLSX summary.

Pipeline:
  Stage 1 — Azure Document Intelligence  (prebuilt-layout → markdown)
  Stage 2 — Azure OpenAI o4-mini         (structured extraction via Pydantic)
  Stage 3 — VE3 / Type lookup            (embedded DIC Excel)
  Stage 4 — Renamed PDF bytes + XLSX     (in-memory, downloadable)

Output filename pattern:
    {VE3}-{DDMMYYYY}-{HP|ZP|NP} {Aufzug|Fahrtreppe} {Fabriknummer}.pdf
"""

import hashlib
import io
import logging
import os
import re
import zipfile
from datetime import date

import pandas as pd
import streamlit as st
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.core.credentials import AzureKeyCredential
from openai import AzureOpenAI
from pydantic import BaseModel, Field

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)-7s  %(message)s")
logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="ZÜS PDF Renamer",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
#  CSS  —  Figma-inspired, consistent with the lift components project
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    * { font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; }

    .main { background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); padding: 2rem; }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1a1f36 0%, #0f1419 100%);
    }

    h1, h2, h3 { font-weight: 600; letter-spacing: -0.02em; color: #1a1f36; }
    h1 { font-size: 2.5rem; margin-bottom: 0.5rem; }
    h2 { font-size: 1.75rem; margin-top: 2rem; }
    h3 { font-size: 1.25rem; margin-top: 1.5rem; }

    /* Cards */
    .analysis-card {
        background: white;
        border-radius: 16px;
        padding: 28px;
        margin: 16px 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05), 0 1px 3px rgba(0,0,0,0.1);
        border: 1px solid #e5e7eb;
        transition: all 0.3s ease;
    }
    .analysis-card:hover {
        box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04);
        transform: translateY(-2px);
    }

    /* Hero */
    .hero-title {
        font-size: 3rem;
        font-weight: 700;
        background: linear-gradient(135deg, #6366f1 0%, #4f46e5 50%, #4338ca 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-bottom: 0.5rem;
        line-height: 1.2;
    }
    .hero-subtitle {
        font-size: 1.1rem;
        color: #6b7280;
        margin-bottom: 2rem;
        font-weight: 400;
        line-height: 1.6;
    }

    /* Status badges */
    .status-ok {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white; padding: 6px 14px; border-radius: 8px;
        font-weight: 600; font-size: 0.8rem; display: inline-block; margin: 3px 0;
    }
    .status-warn {
        background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        color: white; padding: 6px 14px; border-radius: 8px;
        font-weight: 600; font-size: 0.8rem; display: inline-block; margin: 3px 0;
    }
    .status-err {
        background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
        color: white; padding: 6px 14px; border-radius: 8px;
        font-weight: 600; font-size: 0.8rem; display: inline-block; margin: 3px 0;
    }

    /* Metric cards */
    .metric-card {
        background: white; border-radius: 14px; padding: 24px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.07);
        border: 1px solid #e5e7eb;
    }
    .metric-number { font-size: 2.8rem; font-weight: 700; line-height: 1; }
    .metric-label  { font-size: 0.85rem; color: #6b7280; margin-top: 6px; font-weight: 500; }

    /* Filename pill */
    .filename-tag {
        background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%);
        color: white; padding: 4px 12px; border-radius: 20px;
        font-size: 0.78rem; font-weight: 500; display: inline-block;
        margin: 3px 2px; font-family: monospace;
    }
    .new-filename-tag {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white; padding: 4px 12px; border-radius: 20px;
        font-size: 0.78rem; font-weight: 500; display: inline-block;
        margin: 3px 2px; font-family: monospace;
    }

    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%);
        color: white; border: none; border-radius: 10px;
        font-weight: 600; font-size: 0.95rem;
        transition: all 0.2s ease;
        box-shadow: 0 4px 6px rgba(99,102,241,0.25);
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 14px rgba(99,102,241,0.38);
    }

    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, #6366f1 0%, #4f46e5 100%);
        border-radius: 4px;
    }

    /* File uploader */
    [data-testid="stFileUploader"] {
        background: white; border-radius: 12px; padding: 8px;
        border: 2px dashed #6366f1;
    }

    /* Dataframe */
    .stDataFrame { border-radius: 12px; overflow: hidden; }

    /* Scrollbar */
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 3px; }
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #6366f1, #4f46e5);
        border-radius: 3px;
    }

    /* Sidebar text */
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stMarkdown { color: #e2e8f0 !important; }

    /* Divider */
    hr { border-color: #e5e7eb; margin: 1.5rem 0; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  ADMIN AUTH
# ══════════════════════════════════════════════════════════════════════════════
ADMIN_USERNAME = "amr"
ADMIN_PASSWORD_HASH = hashlib.sha256("Micro123456!@#".encode()).hexdigest()


def check_admin(username: str, password: str) -> bool:
    return (
        username == ADMIN_USERNAME
        and hashlib.sha256(password.encode()).hexdigest() == ADMIN_PASSWORD_HASH
    )


# ══════════════════════════════════════════════════════════════════════════════
#  EMBEDDED DIC DATA  (auto-loads from Excel in same directory)
# ══════════════════════════════════════════════════════════════════════════════
_FAHRTREPPE_TYPES = {"Fahrtreppe", "Fahrsteig"}


@st.cache_data(show_spinner=False)
def load_dic_data() -> tuple[dict, dict, int]:
    """Load embedded DIC Excel and build VE3 + type lookup dicts."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    dic_path = os.path.join(script_dir, "DIC (ve3+type).xlsx")
    df = pd.read_excel(dic_path)

    df["_type"] = df["MDESIT_NAME_DE"].apply(
        lambda x: "Fahrtreppe" if x in _FAHRTREPPE_TYPES else "Aufzug"
    )
    ve3_dict = {
        str(k): str(int(v))
        for k, v in zip(df["MDSIT_NUMBER"], df["MDSIT_VE3"])
        if pd.notna(v)
    }
    type_dict = {str(k): v for k, v in zip(df["MDSIT_NUMBER"], df["_type"])}
    return ve3_dict, type_dict, len(df)


# ══════════════════════════════════════════════════════════════════════════════
#  EXAM TYPE MAPPING  —  NP must be checked before HP (critical!)
# ══════════════════════════════════════════════════════════════════════════════
_EXAM_PRIORITY = [
    ("NP", ["nachprüfung"]),
    ("ZP", ["zwischenprüfung"]),
    ("HP", ["hauptprüfung", "wdkprüfung", "wiederkehrende prüfung"]),
]


def map_exam_type(raw: str | None) -> str:
    if not raw:
        return "EXAM_NOT_FOUND"
    lower = raw.strip().lower()
    for code, keywords in _EXAM_PRIORITY:
        if any(kw in lower for kw in keywords):
            return code
    return raw.strip()


# ══════════════════════════════════════════════════════════════════════════════
#  PYDANTIC SCHEMA
# ══════════════════════════════════════════════════════════════════════════════
class ZUSRecord(BaseModel):
    fabriknummer: str | None = Field(
        None,
        description=(
            "Factory / serial number of the lift. "
            "Labels: 'Fabrik-Nr.', 'Fabriknummer', 'Fabrik-Nummer', "
            "'Herstell-Nr.', 'Fabriksnummer', 'Anlagen-Nr.'. "
            "May contain digits, letters, slashes, hyphens, or dots — "
            "e.g. '74/2002', '47NG4659', '93.1.155', '2296/2884'. "
            "Return the value exactly as it appears after reconstruction."
        ),
    )
    pruefungsdatum: date | None = Field(
        None,
        description=(
            "Date the inspection was carried out. "
            "Labels: 'Datum der Prüfung', 'Tag der Prüfung', 'Prüfdatum', 'Prüfungsdatum'. "
            "German format input DD.MM.YYYY — output as ISO 8601 YYYY-MM-DD."
        ),
    )
    pruefungsart_raw: str | None = Field(
        None,
        description=(
            "Full raw text of the inspection type as it appears in the document. "
            "Do NOT abbreviate or map to codes. Return the exact phrase, e.g.: "
            "'Hauptprüfung', 'Zwischenprüfung', "
            "'Wiederkehrende Prüfung (Hauptprüfung)', "
            "'Wiederkehrende Prüfung (Hauptprüfung) - Nachprüfung', "
            "'Nicht durchgeführte Hauptprüfung'."
        ),
    )
    zus_unternehmen: str | None = Field(
        None,
        description=(
            "Name of the ZÜS inspection company. Identify from website URLs, "
            "email domains, logos, letterheads, or signatures in the document. "
            "Known companies (return exactly as listed): "
            "DEKRA, GTÜ, SGS-TÜV Saar, TÜV Austria Deutschland, "
            "TÜV Hessen, TÜV Nord, TÜV Rheinland, TÜV Süd, TÜV Thüringen."
        ),
    )


# ══════════════════════════════════════════════════════════════════════════════
#  EXTRACTION PROMPT
# ══════════════════════════════════════════════════════════════════════════════
ZUS_PROMPT = """\
You are a specialized extraction engine for ZÜS inspection reports (Prüfbescheinigungen)
for lifts and escalators processed by a German lift management company.

## Input
You receive markdown text converted from a PDF by Azure Document Intelligence layout analysis.
The source may be scanned or digital. Treat all input as potentially degraded:
- Characters, words, or tokens may be split, merged, or transposed by OCR
- Table structures may be flattened or columns misaligned
- German compound words are especially prone to mid-word line breaks
- Umlauts (ä, ö, ü, ß) may be mangled

## Your Core Responsibility
Reconstruct and extract four fields from the document. Before emitting any value,
ask: "Does this look like a real serial number, date, inspection label, or company
name in a German ZÜS context?" If not, reconstruct it using surrounding context.

## Reconstruction Examples
| Raw (degraded)               | Extracted (corrected)  | Issue                  |
|------------------------------|------------------------|------------------------|
| "Fabrik- Nr.: 74 / 2002"     | "74/2002"              | label fused, spaces    |
| "Anlagen-Nr 4 712 3"         | "47123"                | digits split by spaces |
| "Prüfungsdaturn: 03.01.2O25" | 2025-01-03             | OCR noise rn→m, 0→O   |
| "Haupt prüfung"              | "Hauptprüfung"         | compound word split    |
| "TUV Nord"                   | "TÜV Nord"             | umlaut mangled         |

## Field-specific Rules

**fabriknummer**
- Look for: Fabrik-Nr., Fabriknummer, Herstell-Nr., Fabriksnummer, Anlagen-Nr.
- May be alphanumeric with slashes, hyphens, dots: 74/2002, 47NG4659, 93.1.155
- Strip leading/trailing whitespace but preserve internal separators exactly

**pruefungsdatum**
- Look for: Datum der Prüfung, Tag der Prüfung, Prüfdatum, Prüfungsdatum
- German input: DD.MM.YYYY → output: YYYY-MM-DD (ISO 8601)

**pruefungsart_raw**
- Return the FULL raw phrase as it appears in the document after reconstruction
- Do NOT map to codes. Examples of valid values:
    "Hauptprüfung"
    "Zwischenprüfung"
    "Wiederkehrende Prüfung (Hauptprüfung)"
    "Wiederkehrende Prüfung (Hauptprüfung) - Nachprüfung"
    "Wiederkehrende Prüfung (Zwischenprüfung) - Nachprüfung"
    "Nicht durchgeführte Hauptprüfung"
- Look near: "Prüfbescheinigung", "Aufzeichnung des Prüfergebnisses",
    "Art der Prüfung", "über die", "zur", "gemäß §"

**zus_unternehmen**
- Identify from website URLs (www.dekra.com, tuev-nord.de …),
    email domains (@tuvsud.com …), logos, or letterheads
- Return EXACTLY one of these standard names:
    DEKRA | GTÜ | SGS-TÜV Saar | TÜV Austria Deutschland |
    TÜV Hessen | TÜV Nord | TÜV Rheinland | TÜV Süd | TÜV Thüringen
- If none match, return the company name as it appears in the document

## Output
Return ONLY a valid JSON object matching the provided schema.
No prose, no explanation, no markdown fences.
"""


# ══════════════════════════════════════════════════════════════════════════════
#  STAGE 1 — Azure Document Intelligence  (sync)
# ══════════════════════════════════════════════════════════════════════════════
def analyze_pdf_bytes(client: DocumentIntelligenceClient, file_bytes: bytes) -> str:
    """Run ADI prebuilt-layout on raw bytes and return markdown content."""
    poller = client.begin_analyze_document(
        "prebuilt-layout",
        io.BytesIO(file_bytes),
        output_content_format="markdown",
    )
    result = poller.result()
    return result.content or ""


# ══════════════════════════════════════════════════════════════════════════════
#  STAGE 2 — Azure OpenAI structured extraction  (sync)
# ══════════════════════════════════════════════════════════════════════════════
def extract_record(client: AzureOpenAI, markdown: str, deploy: str) -> ZUSRecord:
    """Extract ZUSRecord from ADI markdown using o4-mini structured outputs."""
    response = client.beta.chat.completions.parse(
        model=deploy,
        messages=[
            {"role": "system", "content": ZUS_PROMPT},
            {"role": "user",   "content": markdown},
        ],
        response_format=ZUSRecord,
        max_completion_tokens=1000,
    )
    return response.choices[0].message.parsed


# ══════════════════════════════════════════════════════════════════════════════
#  STAGE 3 — VE3 / Type lookup
# ══════════════════════════════════════════════════════════════════════════════
def lookup_ve3_type(
    fabriknummer: str | None,
    ve3_dict: dict,
    type_dict: dict,
) -> tuple[str, str]:
    """Multi-strategy Fabriknummer → (VE3, type) lookup."""
    if not fabriknummer:
        return "VE3_NOT_FOUND", "TYPE_NOT_FOUND"

    candidates = [
        fabriknummer,
        re.sub(r"\s*([./-])\s*", r"\1", fabriknummer),
        fabriknummer.replace("_", "/"),
        re.sub(r"[^0-9]", "", fabriknummer),
        fabriknummer.split("/")[0].strip() if "/" in fabriknummer else None,
    ]
    for key in candidates:
        if key and key in ve3_dict:
            return ve3_dict[key], type_dict.get(key, "TYPE_NOT_FOUND")
    return "VE3_NOT_FOUND", "TYPE_NOT_FOUND"


# ══════════════════════════════════════════════════════════════════════════════
#  FILENAME HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def sanitize(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "", name)
    return re.sub(r"\s+", " ", name).strip()


def unique_name(existing: set, filename: str) -> str:
    if filename not in existing:
        return filename
    stem, ext = os.path.splitext(filename)
    counter = 1
    while f"{stem}_{counter}{ext}" in existing:
        counter += 1
    return f"{stem}_{counter}{ext}"


# ══════════════════════════════════════════════════════════════════════════════
#  STAGE 4 — Full pipeline for a single PDF
# ══════════════════════════════════════════════════════════════════════════════
_EMPTY_ROW = {
    "original_name": "", "new_name": "", "zus_unternehmen": "",
    "fabriknummer": "", "ve3": "", "type": "", "datum": "",
    "pruefungsart_raw": "", "exam_code": "", "status": "", "pdf_bytes": None,
}


def process_pdf(
    filename: str,
    file_bytes: bytes,
    adi_client: DocumentIntelligenceClient,
    oai_client: AzureOpenAI,
    deploy: str,
    ve3_dict: dict,
    type_dict: dict,
    existing_names: set,
) -> dict:
    """Run the 4-stage pipeline for one PDF. Returns a result row dict."""
    row = {**_EMPTY_ROW, "original_name": filename}

    # Stage 1: ADI → markdown
    try:
        markdown = analyze_pdf_bytes(adi_client, file_bytes)
    except Exception as exc:
        logger.error("ADI failed for %s: %s", filename, exc)
        row["status"] = f"ADI_ERROR – {exc}"
        return row

    if not markdown.strip():
        row["status"] = "SKIP – empty ADI output"
        return row

    # Stage 2: OpenAI extraction
    try:
        record = extract_record(oai_client, markdown, deploy)
    except Exception as exc:
        logger.error("OpenAI extraction failed for %s: %s", filename, exc)
        row["status"] = f"OAI_ERROR – {exc}"
        return row

    # Stage 3: VE3 + type lookup
    ve3, typ = lookup_ve3_type(record.fabriknummer, ve3_dict, type_dict)
    exam_code    = map_exam_type(record.pruefungsart_raw)
    date_fn      = record.pruefungsdatum.strftime("%d%m%Y")   if record.pruefungsdatum else "DATE_NOT_FOUND"
    date_display = record.pruefungsdatum.strftime("%d.%m.%Y") if record.pruefungsdatum else "DATE_NOT_FOUND"
    fabrik       = record.fabriknummer or "FABRIK_NOT_FOUND"

    # Stage 4: build unique filename + store bytes
    new_name = sanitize(f"{ve3}-{date_fn}-{exam_code} {typ} {fabrik}.pdf")
    new_name = unique_name(existing_names, new_name)
    existing_names.add(new_name)

    logger.info(
        "OK  %s  →  %s  [ZÜS: %s | VE3: %s | %s | %s]",
        filename, new_name, record.zus_unternehmen, ve3, typ, exam_code,
    )

    return {
        "original_name":   filename,
        "new_name":        new_name,
        "zus_unternehmen": record.zus_unternehmen or "",
        "fabriknummer":    fabrik,
        "ve3":             ve3,
        "type":            typ,
        "datum":           date_display,
        "pruefungsart_raw": record.pruefungsart_raw or "",
        "exam_code":       exam_code,
        "status":          "OK",
        "pdf_bytes":       file_bytes,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  CREDENTIAL HELPERS
# ══════════════════════════════════════════════════════════════════════════════
_CRED_KEYS = ["di_endpoint", "di_key", "oai_endpoint", "oai_key", "oai_deploy", "oai_version"]
_SECRET_MAP = {
    "di_endpoint":  "AZURE_DI_ENDPOINT",
    "di_key":       "AZURE_DI_KEY",
    "oai_endpoint": "AZURE_OAI_ENDPOINT",
    "oai_key":      "AZURE_OAI_KEY",
    "oai_deploy":   "AZURE_OAI_DEPLOY",
    "oai_version":  "AZURE_OAI_VERSION",
}
_DEFAULTS = {
    "oai_deploy":  "o4-mini",
    "oai_version": "2024-12-01-preview",
}


def get_credentials() -> dict:
    """Return merged credentials: st.secrets → session state overrides."""
    creds = {}
    for key, secret_name in _SECRET_MAP.items():
        try:
            val = st.secrets.get(secret_name, "")
            if val:
                creds[key] = val
        except Exception:
            pass

    for key in _CRED_KEYS:
        ss_val = st.session_state.get(f"cred_{key}", "")
        if ss_val:
            creds[key] = ss_val

    for key, default in _DEFAULTS.items():
        creds.setdefault(key, default)

    return creds


def credentials_valid(creds: dict) -> bool:
    return all(creds.get(k) for k in ["di_endpoint", "di_key", "oai_endpoint", "oai_key"])


# ══════════════════════════════════════════════════════════════════════════════
#  DOWNLOAD HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def build_zip(ok_results: list[dict]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in ok_results:
            if r.get("pdf_bytes"):
                zf.writestr(r["new_name"], r["pdf_bytes"])
    return buf.getvalue()


def build_xlsx(results: list[dict]) -> bytes:
    cols = ["original_name", "new_name", "zus_unternehmen", "fabriknummer",
            "ve3", "type", "datum", "pruefungsart_raw", "exam_code", "status"]
    df = pd.DataFrame([{c: r.get(c, "") for c in cols} for r in results])
    df.columns = ["Original Name", "New Name", "ZÜS Company", "Fabrik-Nr.",
                  "VE3", "Type", "Date", "Prüfungsart (raw)", "Exam Code", "Status"]
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  SESSION STATE INIT
# ══════════════════════════════════════════════════════════════════════════════
def init_session_state() -> None:
    defaults = {
        "admin_authenticated": False,
        "results": [],
        "processing_done": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main() -> None:
    init_session_state()

    # Load embedded DIC data
    dic_ok = False
    ve3_dict, type_dict, dic_count = {}, {}, 0
    try:
        ve3_dict, type_dict, dic_count = load_dic_data()
        dic_ok = True
    except Exception as exc:
        logger.error("DIC load failed: %s", exc)

    creds   = get_credentials()
    creds_ok = credentials_valid(creds)

    # ── SIDEBAR ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(
            "<h2 style='color:#e2e8f0;margin-top:0;'>📄 ZÜS PDF Renamer</h2>",
            unsafe_allow_html=True,
        )
        st.markdown(
            "<p style='color:#94a3b8;font-size:0.875rem;margin-top:-8px;'>"
            "Automated inspection report pipeline</p>",
            unsafe_allow_html=True,
        )
        st.divider()

        # System status
        st.markdown("<p style='color:#cbd5e1;font-weight:600;margin-bottom:8px;'>System Status</p>", unsafe_allow_html=True)

        if dic_ok:
            st.markdown(f'<div class="status-ok">✓ DIC Database ({dic_count:,} records)</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-err">✗ DIC Database — file missing</div>', unsafe_allow_html=True)

        if creds_ok:
            st.markdown('<div class="status-ok">✓ Azure Credentials</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-warn">⚠ Credentials not set</div>', unsafe_allow_html=True)

        st.divider()

        # Admin panel
        with st.expander("🔐 Admin Configuration"):
            if not st.session_state.admin_authenticated:
                st.markdown(
                    '<div style="background:#f59e0b;color:white;padding:8px 12px;'
                    'border-radius:8px;font-size:0.8rem;font-weight:600;margin-bottom:12px;">'
                    "⚠ Admin credentials required</div>",
                    unsafe_allow_html=True,
                )
                username = st.text_input("Username", key="admin_user")
                password = st.text_input("Password", type="password", key="admin_pass")
                if st.button("🔓 Verify", use_container_width=True):
                    if check_admin(username, password):
                        st.session_state.admin_authenticated = True
                        st.rerun()
                    else:
                        st.error("❌ Invalid credentials")

            else:
                st.success("🔓 Admin Access Granted")

                tab_di, tab_oai = st.tabs(["🔵 Azure DI", "🟢 Azure OpenAI"])

                with tab_di:
                    st.text_input(
                        "DI Endpoint",
                        value=creds.get("di_endpoint", ""),
                        key="cred_di_endpoint",
                        placeholder="https://…cognitiveservices.azure.com/",
                    )
                    st.text_input(
                        "DI Key",
                        value=creds.get("di_key", ""),
                        key="cred_di_key",
                        type="password",
                    )

                with tab_oai:
                    st.text_input(
                        "OAI Endpoint",
                        value=creds.get("oai_endpoint", ""),
                        key="cred_oai_endpoint",
                        placeholder="https://…openai.azure.com/",
                    )
                    st.text_input(
                        "OAI Key",
                        value=creds.get("oai_key", ""),
                        key="cred_oai_key",
                        type="password",
                    )
                    st.text_input(
                        "Deployment name",
                        value=creds.get("oai_deploy", "o4-mini"),
                        key="cred_oai_deploy",
                    )
                    st.text_input(
                        "API Version",
                        value=creds.get("oai_version", "2024-12-01-preview"),
                        key="cred_oai_version",
                    )

                st.divider()
                if st.button("🔒 Lock Admin", use_container_width=True):
                    st.session_state.admin_authenticated = False
                    st.rerun()

    # ── HERO ─────────────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style="text-align:center;padding:2.5rem 0 1.5rem 0;">
            <div class="hero-title">ZÜS PDF Renamer</div>
            <div class="hero-subtitle">
                Upload ZÜS inspection reports and get them automatically renamed<br>
                using Azure Document Intelligence + Azure OpenAI o4-mini.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not creds_ok:
        st.info(
            "ℹ️ Azure credentials are not configured. "
            "Open **🔐 Admin Configuration** in the sidebar to enter them, "
            "or add them to `.streamlit/secrets.toml`."
        )

    # ── FILE UPLOADER ─────────────────────────────────────────────────────────
    st.markdown('<div class="analysis-card">', unsafe_allow_html=True)
    st.markdown("### 📂 Upload ZÜS Inspection PDFs")
    st.markdown(
        "<p style='color:#6b7280;font-size:0.9rem;margin-top:-8px;'>"
        "Drag & drop one or more PDF files. Each file is processed independently through the pipeline.</p>",
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        "Upload PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if uploaded_files:
        st.markdown(
            f"<p style='color:#374151;font-weight:600;margin-top:12px;'>"
            f"{len(uploaded_files)} file(s) selected:</p>",
            unsafe_allow_html=True,
        )
        pills = " ".join(
            f'<span class="filename-tag">📄 {f.name}</span>' for f in uploaded_files
        )
        st.markdown(pills, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    # ── PROCESS BUTTON ────────────────────────────────────────────────────────
    col_l, col_c, col_r = st.columns([1, 2, 1])
    with col_c:
        process_clicked = st.button(
            "🚀  Process & Rename PDFs",
            use_container_width=True,
            type="primary",
        )

    if process_clicked:
        if not uploaded_files:
            st.error("Please upload at least one PDF file.")
        elif not creds_ok:
            st.error("Azure credentials are missing. Configure them in the Admin panel.")
        elif not dic_ok:
            st.error("DIC database could not be loaded. Ensure 'DIC (ve3+type).xlsx' is in the app folder.")
        else:
            creds = get_credentials()
            adi_client = DocumentIntelligenceClient(
                endpoint=creds["di_endpoint"],
                credential=AzureKeyCredential(creds["di_key"]),
            )
            oai_client = AzureOpenAI(
                azure_endpoint=creds["oai_endpoint"],
                api_key=creds["oai_key"],
                api_version=creds["oai_version"],
            )
            deploy = creds["oai_deploy"]

            st.session_state.results = []
            existing_names: set = set()

            progress_bar    = st.progress(0.0, text="Initialising pipeline…")
            status_slot     = st.empty()
            live_table_slot = st.empty()

            for idx, pdf_file in enumerate(uploaded_files):
                frac = idx / len(uploaded_files)
                progress_bar.progress(frac, text=f"Processing {pdf_file.name}  ({idx + 1}/{len(uploaded_files)})…")
                status_slot.info(f"⏳  Stage 1/4 — Analysing: **{pdf_file.name}**")

                file_bytes = pdf_file.read()
                row = process_pdf(
                    pdf_file.name, file_bytes,
                    adi_client, oai_client, deploy,
                    ve3_dict, type_dict, existing_names,
                )
                st.session_state.results.append(row)

                # Live preview of results so far
                preview_cols = ["original_name", "new_name", "ve3", "type", "datum", "exam_code", "status"]
                df_live = pd.DataFrame(
                    [{c: r.get(c, "") for c in preview_cols} for r in st.session_state.results]
                )
                df_live.columns = ["Original", "New Name", "VE3", "Type", "Date", "Code", "Status"]
                live_table_slot.dataframe(df_live, use_container_width=True)

            progress_bar.progress(1.0, text="Pipeline complete!")
            status_slot.success(f"✅  Finished — {len(uploaded_files)} file(s) processed.")
            st.session_state.processing_done = True

    # ── RESULTS ───────────────────────────────────────────────────────────────
    if st.session_state.results:
        results  = st.session_state.results
        ok_rows  = [r for r in results if r["status"] == "OK"]
        err_rows = [r for r in results if r["status"] != "OK"]

        st.markdown("---")

        # Metric cards
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(
                f'<div class="metric-card"><div class="metric-number" style="color:#6366f1;">'
                f'{len(results)}</div><div class="metric-label">Files Processed</div></div>',
                unsafe_allow_html=True,
            )
        with c2:
            st.markdown(
                f'<div class="metric-card"><div class="metric-number" style="color:#10b981;">'
                f'{len(ok_rows)}</div><div class="metric-label">Successfully Renamed</div></div>',
                unsafe_allow_html=True,
            )
        with c3:
            st.markdown(
                f'<div class="metric-card"><div class="metric-number" style="color:#ef4444;">'
                f'{len(err_rows)}</div><div class="metric-label">Errors / Skipped</div></div>',
                unsafe_allow_html=True,
            )

        # Rename mapping
        if ok_rows:
            st.markdown("---")
            st.markdown("### ✅ Renamed Files")
            for r in ok_rows:
                st.markdown(
                    f'<span class="filename-tag">📄 {r["original_name"]}</span>'
                    f' &nbsp;→&nbsp; '
                    f'<span class="new-filename-tag">✓ {r["new_name"]}</span>',
                    unsafe_allow_html=True,
                )

        # Error details
        if err_rows:
            st.markdown("---")
            with st.expander(f"⚠️ {len(err_rows)} file(s) with errors — click to expand"):
                for r in err_rows:
                    st.markdown(f"**{r['original_name']}** — `{r['status']}`")

        # Full results table
        st.markdown("---")
        st.markdown("### 📊 Full Results Table")
        display_cols = [
            "original_name", "new_name", "zus_unternehmen", "fabriknummer",
            "ve3", "type", "datum", "pruefungsart_raw", "exam_code", "status",
        ]
        df_full = pd.DataFrame([{c: r.get(c, "") for c in display_cols} for r in results])
        df_full.columns = [
            "Original Name", "New Name", "ZÜS Company", "Fabrik-Nr.",
            "VE3", "Type", "Date", "Prüfungsart (raw)", "Exam Code", "Status",
        ]
        st.dataframe(df_full, use_container_width=True, height=320)

        # Downloads
        st.markdown("---")
        st.markdown("### 📥 Downloads")
        dl1, dl2 = st.columns(2)

        with dl1:
            if ok_rows:
                zip_bytes = build_zip(ok_rows)
                st.download_button(
                    label="📦  Download Renamed PDFs  (ZIP)",
                    data=zip_bytes,
                    file_name="renamed_pdfs.zip",
                    mime="application/zip",
                    use_container_width=True,
                )
            else:
                st.info("No successfully renamed files to download.")

        with dl2:
            xlsx_bytes = build_xlsx(results)
            st.download_button(
                label="📊  Download Summary  (XLSX)",
                data=xlsx_bytes,
                file_name="zus_pipeline_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # Reset
        st.markdown("---")
        col_l2, col_c2, col_r2 = st.columns([1, 2, 1])
        with col_c2:
            if st.button("🔄  Reset — Process New Files", use_container_width=True):
                st.session_state.results = []
                st.session_state.processing_done = False
                st.rerun()


if __name__ == "__main__":
    main()
