"""
ZÜS PDF Rename Pipeline — Streamlit Web App
============================================
Upload ZÜS inspection PDFs → get automatically renamed files + XLSX summary.

Pipeline:
  Stage 1 — Azure Document Intelligence  (prebuilt-layout → markdown)
  Stage 2 — Azure OpenAI o4-mini         (structured extraction via Pydantic)
  Stage 3 — VE3 / Type lookup            (embedded DIC Excel)
  Stage 4 — Renamed PDF bytes + XLSX     (in-memory, downloadable)

Two tabs, sharing the ADI stage, the Azure credentials and the admin panel:

  "ZUES Renamer"  general pipeline, DIC lookup
      {VE3}-{DDMMYYYY}-{HP|ZP|NP} {Aufzug|Fahrtreppe} {Fabriknummer}.pdf

  "Gesobau_ZUES"  GESOBAU AG only, Gesobau masterlist lookup
      {HP|ZP|NP}-{Strasse}-{Fabriknummer}.pdf
      e.g.  HP-Senftenberger Ring 37-10987474.pdf
"""

import base64
import gzip
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
    initial_sidebar_state="collapsed",
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
#  EMBEDDED LOOKUP TABLES
#  The tables hold customer data, so they are NOT committed to the repository.
#  They live in Streamlit secrets as a gzipped + base64-encoded CSV; locally the
#  original xlsx (gitignored) is used as a fallback so the dev loop is unchanged.
#  Build a blob with:
#      import gzip, base64, pandas as pd
#      csv = pd.read_excel("DIC (ve3+type).xlsx").to_csv(index=False).encode()
#      print(base64.b64encode(gzip.compress(csv, 9)).decode())
# ══════════════════════════════════════════════════════════════════════════════
def _load_table(secret_key: str, xlsx_filename: str) -> pd.DataFrame:
    """Return a lookup table from st.secrets, else from the local xlsx file.

    The secret holds a gzipped, base64-encoded CSV; whitespace and line breaks
    inside it are ignored, so it can be pasted as a wrapped TOML block string.
    Values come back as strings, which every caller already normalises with
    str()/float(), so both paths produce identical lookups.
    """
    blob = ""
    try:
        blob = str(st.secrets.get(secret_key, "") or "")
    except Exception:
        pass

    if blob.strip():
        raw = gzip.decompress(base64.b64decode("".join(blob.split())))
        return pd.read_csv(io.BytesIO(raw), dtype=str)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    return pd.read_excel(os.path.join(script_dir, xlsx_filename))


_FAHRTREPPE_TYPES = {"Fahrtreppe", "Fahrsteig"}


@st.cache_data(show_spinner=False)
def load_dic_data() -> tuple[dict, dict, int]:
    """Load the DIC table and build VE3 + type lookup dicts."""
    df = _load_table("DIC_TABLE_B64", "DIC (ve3+type).xlsx")

    df["_type"] = df["MDESIT_NAME_DE"].apply(
        lambda x: "Fahrtreppe" if x in _FAHRTREPPE_TYPES else "Aufzug"
    )
    ve3_dict = {
        str(k): str(int(float(v)))          # float() so "16332.0" and 16332.0 both work
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
# ██  GESOBAU  ████████████████████████████████████████████████████████████████
#  Second pipeline — same ADI stage, customer-specific extraction + naming.
#  Output filename pattern:  {HP|ZP|NP}-{Straße}-{Fabriknummer}.pdf
#      e.g.  HP-Senftenberger Ring 37-10987474.pdf
# ══════════════════════════════════════════════════════════════════════════════
GESOBAU_XLSX = "Gesobau (fabrik+strasse).xlsx"


def _fab_keys(raw: str) -> list[str]:
    """Candidate lookup keys derived from one Fabriknummer string."""
    raw = raw.strip()
    out = [
        raw,
        re.sub(r"\s*([./-])\s*", r"\1", raw),   # "74 / 2002"  -> "74/2002"
        raw.replace(" ", ""),                   # "4 712 3"    -> "47123"
        raw.replace("_", "/"),
    ]
    seen, keys = set(), []
    for k in out:
        if k and k not in seen:
            seen.add(k)
            keys.append(k)
    return keys


@st.cache_data(show_spinner=False)
def load_gesobau_data() -> tuple[dict, dict, dict, int]:
    """Load the embedded Gesobau masterlist -> exact / upper / digits-only lookups.

    Verified against the current export: no Fabriknummer maps to two different
    streets under any of the three keyings, so the fallbacks cannot introduce
    ambiguity. Rows without a Fabriknummer or Straße are dropped — that also
    removes the trailing blank row and the "Applied filters:" footer row the
    source system appends to every export.
    """
    df = _load_table("GESOBAU_TABLE_B64", GESOBAU_XLSX)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(subset=["Fabriknummer", "Straße"])

    exact: dict[str, str] = {}
    upper: dict[str, str] = {}
    digits: dict[str, str] = {}
    for fab_raw, street_raw in zip(df["Fabriknummer"], df["Straße"]):
        fab, street = str(fab_raw).strip(), str(street_raw).strip()
        if not fab or not street or fab.lower() == "nan":
            continue
        exact.setdefault(fab, street)
        upper.setdefault(fab.upper(), street)
        only_digits = re.sub(r"\D", "", fab)
        if only_digits:
            digits.setdefault(only_digits, street)
    return exact, upper, digits, len(exact)


def lookup_gesobau_street(
    fabriknummer: str | None,
    exact: dict,
    upper: dict,
    digits: dict,
) -> str | None:
    """Multi-strategy Fabriknummer -> Straße lookup. None when not in the list."""
    if not fabriknummer:
        return None

    keys = _fab_keys(fabriknummer)
    for key in keys:                                # 1 — verbatim
        if key in exact:
            return exact[key]
    for key in keys:                                # 2 — case-insensitive
        if key.upper() in upper:
            return upper[key.upper()]

    only_digits = re.sub(r"\D", "", fabriknummer)   # 3 — prefix / punctuation drift
    if only_digits and only_digits in digits:       #     (B200451 <-> 200451)
        return digits[only_digits]
    return None


# ── Gesobau extraction schema ────────────────────────────────────────────────
class GesobauRecord(BaseModel):
    fabriknummer: str | None = Field(
        None,
        description=(
            "Factory / serial number of the lift, from the field labelled 'Fabriknr.', "
            "'Fabrik-Nr.', 'Fabriknummer', 'Herstell-Nr.' or 'Anlagen-Nr.' inside the "
            "'Allgemeine Daten zur Anlage' block. May be digits only ('10987474', '200451') "
            "or alphanumeric ('47NBR275', '6KF61010', 'B200451', '20041956-1'). "
            "Return it exactly as printed — keep letters, leading zeros and internal separators."
        ),
    )
    pruefungsart_raw: str | None = Field(
        None,
        description=(
            "The FULL raw text of the field labelled 'Art der Prüfung' (the heading may read "
            "'Art der Prüfung/ Prüfgrundlage'). Return only that field's value, e.g. "
            "'Hauptprüfung inkl. Ersatzsystem', 'Zwischenprüfung', 'Nachprüfung', "
            "'Wiederkehrende Prüfung (Hauptprüfung) - Nachprüfung'. "
            "CRITICAL: ignore sentences elsewhere in the document such as 'Eine Nachprüfung "
            "ist bis zum TT.MM.JJJJ erforderlich' in the Prüfergebnis — that is a future "
            "requirement, not the type of THIS inspection. Do not abbreviate or map to a code."
        ),
    )
    standort_strasse: str | None = Field(
        None,
        description=(
            "Street name and house number of the lift LOCATION, from the 'Standort' field "
            "('Senftenberger Ring 37, 13435 Berlin' -> 'Senftenberger Ring 37'). "
            "CRITICAL: this is NOT the 'Betreiber' / 'Auftraggeber' address — that is the "
            "GESOBAU AG head office ('Stiftsweg 1, 13187 Berlin') and must never be returned. "
            "Exclude postal code, city and any 'GESOBAU AG WHG ###' object label. Keep "
            "house-number suffix letters ('37', '44f', '42Q'). Null if no Standort is present."
        ),
    )
    pruefdatum: date | None = Field(
        None,
        description=(
            "Date the inspection was carried out ('Prüfdatum', 'Datum der Prüfung', "
            "'Tag der Prüfung'). German DD.MM.YYYY input -> ISO 8601 YYYY-MM-DD output."
        ),
    )
    zus_unternehmen: str | None = Field(
        None,
        description=(
            "Name of the ZÜS inspection company — from logo, letterhead, website URL or "
            "e-mail domain. Usually 'TÜV Thüringen' for this customer. Return exactly one of: "
            "DEKRA | GTÜ | SGS-TÜV Saar | TÜV Austria Deutschland | TÜV Hessen | TÜV Nord | "
            "TÜV Rheinland | TÜV Süd | TÜV Thüringen. If none match, return it as printed."
        ),
    )


GESOBAU_PROMPT = """\
You are a specialized extraction engine for ZÜS inspection reports (Prüfbescheinigungen)
for lifts operated by the Berlin housing company GESOBAU AG. Reports come mostly from
TÜV Thüringen, but other ZÜS may appear — handle any layout.

## Input
Markdown produced by Azure Document Intelligence layout analysis of a PDF. The source may be
scanned or digital. Treat all input as potentially degraded:
- Characters, words or tokens may be split, merged or transposed by OCR
- Table structures may be flattened or columns misaligned
- German compound words are especially prone to mid-word line breaks
- Umlauts (ä, ö, ü, ß) may be mangled

## Reconstruction examples
| Raw (degraded)                | Extracted                | Issue                 |
|-------------------------------|--------------------------|-----------------------|
| "Fabrik- Nr.: 74 / 2002"      | "74/2002"                | label fused, spaces   |
| "Anlagen-Nr 4 712 3"          | "47123"                  | digits split          |
| "Prüfungsdaturn: 03.01.2O25"  | 2025-01-03               | OCR noise rn->m, O->0 |
| "Haupt prüfung"               | "Hauptprüfung"           | compound word split   |
| "Senftenberger Ring  3 7"     | "Senftenberger Ring 37"  | digits split          |
| "Wilhelmsruher Darnm 141"     | "Wilhelmsruher Damm 141" | OCR noise rn->m       |

## Two traps you MUST avoid

**1. Address trap.** These documents carry TWO addresses.
   - 'Auftraggeber' / 'Betreiber' -> GESOBAU AG head office, *Stiftsweg 1, 13187 Berlin*
   - 'Standort' -> the building the lift actually stands in
   Always return the Standort street. Never return Stiftsweg 1.

**2. Nachprüfung trap.** The 'Prüfergebnis' section frequently contains
   "Eine Nachprüfung ist bis zum TT.MM.JJJJ erforderlich". That is a *future* follow-up
   requirement, not this report's own type. The inspection type is ONLY the value printed
   under the label 'Art der Prüfung'.

## Output
Return ONLY a valid JSON object matching the provided schema.
No prose, no explanation, no markdown fences.
"""


_GESOBAU_EMPTY_ROW = {
    "original_name": "", "new_name": "", "exam_code": "", "strasse": "",
    "fabriknummer": "", "datum": "", "pruefungsart_raw": "",
    "zus_unternehmen": "", "status": "", "pdf_bytes": None,
}


def extract_gesobau_record(client: AzureOpenAI, markdown: str, deploy: str) -> GesobauRecord:
    """Extract a GesobauRecord from ADI markdown using structured outputs."""
    response = client.beta.chat.completions.parse(
        model=deploy,
        messages=[
            {"role": "system", "content": GESOBAU_PROMPT},
            {"role": "user",   "content": markdown},
        ],
        response_format=GesobauRecord,
        max_completion_tokens=2000,
    )
    return response.choices[0].message.parsed


def process_pdf_gesobau(
    filename: str,
    file_bytes: bytes,
    adi_client: DocumentIntelligenceClient,
    oai_client: AzureOpenAI,
    deploy: str,
    exact: dict,
    upper: dict,
    digits: dict,
    existing_names: set,
) -> dict:
    """Run the Gesobau pipeline for one PDF. Returns a result row dict."""
    row = {**_GESOBAU_EMPTY_ROW, "original_name": filename}

    # Stage 1 — ADI -> markdown
    try:
        markdown = analyze_pdf_bytes(adi_client, file_bytes)
    except Exception as exc:
        logger.error("ADI failed for %s: %s", filename, exc)
        row["status"] = f"ADI_ERROR – {exc}"
        return row

    if not markdown.strip():
        row["status"] = "SKIP – empty ADI output"
        return row

    # Stage 2 — structured extraction
    try:
        record = extract_gesobau_record(oai_client, markdown, deploy)
    except Exception as exc:
        logger.error("OpenAI extraction failed for %s: %s", filename, exc)
        row["status"] = f"OAI_ERROR – {exc}"
        return row

    # Stage 3 — Fabriknummer (as printed) -> Straße from the masterlist, falling
    #           back to the Standort street read from the PDF itself.
    fabrik  = (record.fabriknummer or "").strip() or "FABRIK_NOT_FOUND"
    strasse = lookup_gesobau_street(record.fabriknummer, exact, upper, digits)
    if not strasse:
        strasse = (record.standort_strasse or "").strip() or "STRASSE_NOT_FOUND"
    exam_code = map_exam_type(record.pruefungsart_raw)
    datum     = record.pruefdatum.strftime("%d.%m.%Y") if record.pruefdatum else ""

    # Stage 4 — filename + bytes
    new_name = sanitize(f"{exam_code}-{strasse}-{fabrik}.pdf")
    new_name = unique_name(existing_names, new_name)
    existing_names.add(new_name)

    logger.info(
        "OK  %s  ->  %s  [ZÜS: %s | %s | %s]",
        filename, new_name, record.zus_unternehmen, strasse, exam_code,
    )

    return {
        "original_name":    filename,
        "new_name":         new_name,
        "exam_code":        exam_code,
        "strasse":          strasse,
        "fabriknummer":     fabrik,
        "datum":            datum,
        "pruefungsart_raw": record.pruefungsart_raw or "",
        "zus_unternehmen":  record.zus_unternehmen or "",
        "status":           "OK",
        "pdf_bytes":        file_bytes,
    }


GESOBAU_COLS   = ["original_name", "new_name", "exam_code", "strasse", "fabriknummer",
                  "datum", "pruefungsart_raw", "zus_unternehmen", "status"]
GESOBAU_LABELS = ["Original Name", "New Name", "Exam Code", "Straße", "Fabrik-Nr.",
                  "Prüfdatum", "Prüfungsart (raw)", "ZÜS Company", "Status"]


def build_xlsx_gesobau(results: list[dict]) -> bytes:
    df = pd.DataFrame([{c: r.get(c, "") for c in GESOBAU_COLS} for r in results])
    df.columns = GESOBAU_LABELS
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  TAB STYLING
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: transparent;
        border-bottom: 1px solid #e5e7eb;
    }
    .stTabs [data-baseweb="tab"] {
        height: 46px;
        padding: 0 24px;
        border-radius: 10px 10px 0 0;
        background: rgba(255,255,255,0.7);
        border: 1px solid #e5e7eb;
        border-bottom: none;
        font-weight: 600;
        font-size: 0.95rem;
        color: #4b5563;
    }
    .stTabs [data-baseweb="tab"]:hover { background: #ffffff; color: #4f46e5; }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
        color: #ffffff !important;
        border-color: transparent !important;
        box-shadow: 0 4px 10px rgba(99,102,241,0.28);
    }
    .stTabs [data-baseweb="tab-highlight"] { background: transparent; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  SESSION STATE INIT
# ══════════════════════════════════════════════════════════════════════════════
def init_session_state() -> None:
    defaults = {
        "admin_authenticated":     False,
        "results":                 [],
        "processing_done":         False,
        "results_gesobau":         [],
        "processing_done_gesobau": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ══════════════════════════════════════════════════════════════════════════════
#  SHARED UI BUILDING BLOCKS
# ══════════════════════════════════════════════════════════════════════════════
def make_clients(creds: dict) -> tuple[DocumentIntelligenceClient, AzureOpenAI, str]:
    adi_client = DocumentIntelligenceClient(
        endpoint=creds["di_endpoint"],
        credential=AzureKeyCredential(creds["di_key"]),
    )
    oai_client = AzureOpenAI(
        azure_endpoint=creds["oai_endpoint"],
        api_key=creds["oai_key"],
        api_version=creds["oai_version"],
    )
    return adi_client, oai_client, creds["oai_deploy"]


def render_uploader(ns: str, caption: str) -> list:
    """Drag & drop area + selected-file pills. Returns the uploaded file objects."""
    st.markdown("### 📂 Upload ZÜS Inspection PDFs")
    st.markdown(
        f"<p style='color:#6b7280;font-size:0.9rem;margin-top:-8px;'>{caption}</p>",
        unsafe_allow_html=True,
    )
    uploaded_files = st.file_uploader(
        "Upload PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key=f"upload_{ns}",
    )
    if uploaded_files:
        st.markdown(
            f"<p style='color:#374151;font-weight:600;margin-top:12px;'>"
            f"{len(uploaded_files)} file(s) selected:</p>",
            unsafe_allow_html=True,
        )
        st.markdown(
            " ".join(f'<span class="filename-tag">📄 {f.name}</span>' for f in uploaded_files),
            unsafe_allow_html=True,
        )
    return uploaded_files or []


def run_batch(
    uploaded_files: list,
    process_one,
    results_key: str,
    live_cols: list[str],
    live_labels: list[str],
) -> None:
    """Process every uploaded PDF, streaming a live result table as it goes."""
    st.session_state[results_key] = []
    existing_names: set = set()

    progress_bar    = st.progress(0.0, text="Initialising pipeline…")
    status_slot     = st.empty()
    live_table_slot = st.empty()

    total = len(uploaded_files)
    for idx, pdf_file in enumerate(uploaded_files):
        progress_bar.progress(
            idx / total,
            text=f"Processing {pdf_file.name}  ({idx + 1}/{total})…",
        )
        status_slot.info(f"⏳  Processing **{pdf_file.name}**…")

        row = process_one(pdf_file.name, pdf_file.read(), existing_names)
        st.session_state[results_key].append(row)

        df_live = pd.DataFrame(
            [{c: r.get(c, "") for c in live_cols} for r in st.session_state[results_key]]
        )
        df_live.columns = live_labels
        live_table_slot.dataframe(df_live, use_container_width=True)

    progress_bar.progress(1.0, text="Pipeline complete!")
    status_slot.success(f"✅  Finished — {total} file(s) processed.")


def render_results(
    *,
    results: list[dict],
    display_cols: list[str],
    display_labels: list[str],
    xlsx_builder,
    zip_name: str,
    xlsx_name: str,
    results_key: str,
    done_key: str,
    ns: str,
) -> None:
    """Metrics, rename mapping, errors, full table, downloads and reset."""
    ok_rows  = [r for r in results if r["status"] == "OK"]
    err_rows = [r for r in results if r["status"] != "OK"]

    st.markdown("---")

    c1, c2, c3 = st.columns(3)
    for col, value, color, label in (
        (c1, len(results),  "#6366f1", "Files Processed"),
        (c2, len(ok_rows),  "#10b981", "Successfully Renamed"),
        (c3, len(err_rows), "#ef4444", "Errors / Skipped"),
    ):
        with col:
            st.markdown(
                f'<div class="metric-card"><div class="metric-number" style="color:{color};">'
                f'{value}</div><div class="metric-label">{label}</div></div>',
                unsafe_allow_html=True,
            )

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

    if err_rows:
        st.markdown("---")
        with st.expander(f"⚠️ {len(err_rows)} file(s) with errors — click to expand"):
            for r in err_rows:
                st.markdown(f"**{r['original_name']}** — `{r['status']}`")

    st.markdown("---")
    st.markdown("### 📊 Full Results Table")
    df_full = pd.DataFrame([{c: r.get(c, "") for c in display_cols} for r in results])
    df_full.columns = display_labels
    st.dataframe(df_full, use_container_width=True, height=320)

    st.markdown("---")
    st.markdown("### 📥 Downloads")
    dl1, dl2 = st.columns(2)

    with dl1:
        if ok_rows:
            st.download_button(
                label="📦  Download Renamed PDFs  (ZIP)",
                data=build_zip(ok_rows),
                file_name=zip_name,
                mime="application/zip",
                use_container_width=True,
                key=f"dl_zip_{ns}",
            )
        else:
            st.info("No successfully renamed files to download.")

    with dl2:
        st.download_button(
            label="📊  Download Summary  (XLSX)",
            data=xlsx_builder(results),
            file_name=xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"dl_xlsx_{ns}",
        )

    st.markdown("---")
    _, col_c, _ = st.columns([1, 2, 1])
    with col_c:
        if st.button("🔄  Reset — Process New Files", use_container_width=True, key=f"reset_{ns}"):
            st.session_state[results_key] = []
            st.session_state[done_key] = False
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 1 — STANDARD ZÜS RENAMER
# ══════════════════════════════════════════════════════════════════════════════
STD_COLS   = ["original_name", "new_name", "zus_unternehmen", "fabriknummer",
              "ve3", "type", "datum", "pruefungsart_raw", "exam_code", "status"]
STD_LABELS = ["Original Name", "New Name", "ZÜS Company", "Fabrik-Nr.",
              "VE3", "Type", "Date", "Prüfungsart (raw)", "Exam Code", "Status"]


def render_standard_tab(creds: dict, creds_ok: bool, dic_ok: bool,
                        ve3_dict: dict, type_dict: dict) -> None:
    uploaded_files = render_uploader(
        "std",
        "Drag &amp; drop one or more PDF files. Each file is processed independently "
        "through the pipeline.",
    )

    _, col_c, _ = st.columns([1, 2, 1])
    with col_c:
        process_clicked = st.button(
            "🚀  Process & Rename PDFs",
            use_container_width=True,
            type="primary",
            key="process_std",
        )

    if process_clicked:
        if not uploaded_files:
            st.error("Please upload at least one PDF file.")
        elif not creds_ok:
            st.error("Azure credentials are missing. Configure them in the Admin panel.")
        elif not dic_ok:
            st.error("DIC table could not be loaded. Set DIC_TABLE_B64 in secrets, "
                     "or place 'DIC (ve3+type).xlsx' in the app folder.")
        else:
            adi_client, oai_client, deploy = make_clients(get_credentials())
            run_batch(
                uploaded_files,
                lambda name, data, existing: process_pdf(
                    name, data, adi_client, oai_client, deploy,
                    ve3_dict, type_dict, existing,
                ),
                results_key="results",
                live_cols=["original_name", "new_name", "ve3", "type", "datum", "exam_code", "status"],
                live_labels=["Original", "New Name", "VE3", "Type", "Date", "Code", "Status"],
            )
            st.session_state.processing_done = True

    if st.session_state.results:
        render_results(
            results=st.session_state.results,
            display_cols=STD_COLS,
            display_labels=STD_LABELS,
            xlsx_builder=build_xlsx,
            zip_name="renamed_pdfs.zip",
            xlsx_name="zus_pipeline_output.xlsx",
            results_key="results",
            done_key="processing_done",
            ns="std",
        )


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — GESOBAU_ZÜS
# ══════════════════════════════════════════════════════════════════════════════
def render_gesobau_tab(creds: dict, creds_ok: bool, gesobau_ok: bool,
                       exact: dict, upper: dict, digits: dict) -> None:
    uploaded_files = render_uploader(
        "gesobau",
        "Drag &amp; drop one or more GESOBAU inspection PDFs (mostly TÜV Thüringen). "
        "The street is resolved from the Fabriknummer.",
    )

    _, col_c, _ = st.columns([1, 2, 1])
    with col_c:
        process_clicked = st.button(
            "🚀  Process & Rename PDFs",
            use_container_width=True,
            type="primary",
            key="process_gesobau",
        )

    if process_clicked:
        if not uploaded_files:
            st.error("Please upload at least one PDF file.")
        elif not creds_ok:
            st.error("Azure credentials are missing. Configure them in the Admin panel.")
        elif not gesobau_ok:
            st.error("Gesobau masterlist could not be loaded. Set GESOBAU_TABLE_B64 in secrets, "
                     f"or place '{GESOBAU_XLSX}' in the app folder.")
        else:
            adi_client, oai_client, deploy = make_clients(get_credentials())
            run_batch(
                uploaded_files,
                lambda name, data, existing: process_pdf_gesobau(
                    name, data, adi_client, oai_client, deploy,
                    exact, upper, digits, existing,
                ),
                results_key="results_gesobau",
                live_cols=["original_name", "new_name", "exam_code", "strasse", "fabriknummer", "status"],
                live_labels=["Original", "New Name", "Code", "Straße", "Fabrik-Nr.", "Status"],
            )
            st.session_state.processing_done_gesobau = True

    if st.session_state.results_gesobau:
        render_results(
            results=st.session_state.results_gesobau,
            display_cols=GESOBAU_COLS,
            display_labels=GESOBAU_LABELS,
            xlsx_builder=build_xlsx_gesobau,
            zip_name="gesobau_renamed_pdfs.zip",
            xlsx_name="gesobau_pipeline_output.xlsx",
            results_key="results_gesobau",
            done_key="processing_done_gesobau",
            ns="gesobau",
        )


# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
def render_sidebar(creds: dict, creds_ok: bool,
                   dic_ok: bool, dic_count: int,
                   gesobau_ok: bool, gesobau_count: int) -> None:
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

        st.markdown(
            "<p style='color:#cbd5e1;font-weight:600;margin-bottom:8px;'>System Status</p>",
            unsafe_allow_html=True,
        )

        if dic_ok:
            st.markdown(f'<div class="status-ok">✓ DIC Database ({dic_count:,} records)</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-err">✗ DIC Database — not configured</div>',
                        unsafe_allow_html=True)

        if gesobau_ok:
            st.markdown(f'<div class="status-ok">✓ Gesobau Masterlist ({gesobau_count:,} lifts)</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-err">✗ Gesobau Masterlist — not configured</div>',
                        unsafe_allow_html=True)

        if creds_ok:
            st.markdown('<div class="status-ok">✓ Azure Credentials</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-warn">⚠ Credentials not set</div>', unsafe_allow_html=True)

        st.divider()

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


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main() -> None:
    init_session_state()

    # Embedded lookup databases
    dic_ok, ve3_dict, type_dict, dic_count = False, {}, {}, 0
    try:
        ve3_dict, type_dict, dic_count = load_dic_data()
        dic_ok = True
    except Exception as exc:
        logger.error("DIC load failed: %s", exc)

    gesobau_ok, g_exact, g_upper, g_digits, gesobau_count = False, {}, {}, {}, 0
    try:
        g_exact, g_upper, g_digits, gesobau_count = load_gesobau_data()
        gesobau_ok = True
    except Exception as exc:
        logger.error("Gesobau masterlist load failed: %s", exc)

    creds    = get_credentials()
    creds_ok = credentials_valid(creds)

    render_sidebar(creds, creds_ok, dic_ok, dic_count, gesobau_ok, gesobau_count)

    # ── HERO ─────────────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style="text-align:center;padding:2.5rem 0 1rem 0;">
            <div class="hero-title">ZÜS PDF Renamer</div>
            <div class="hero-subtitle">
                Upload ZÜS inspection reports and get them automatically renamed.
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

    tab_std, tab_gesobau = st.tabs(["📄  Branicks_ZÜS", "🏢  Gesobau_ZÜS"])

    with tab_std:
        render_standard_tab(creds, creds_ok, dic_ok, ve3_dict, type_dict)

    with tab_gesobau:
        render_gesobau_tab(creds, creds_ok, gesobau_ok, g_exact, g_upper, g_digits)


if __name__ == "__main__":
    main()
