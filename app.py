import streamlit as st
import pdfplumber
import pandas as pd
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="EQAS Extractor", layout="wide")
st.title("EQAS Auto Extractor")

uploaded_files = st.file_uploader(
    "Upload EQAS PDF(s)", type=["pdf"], accept_multiple_files=True
)


# ─────────────────────────────────────────────
# NORMALISE unicode minus → ASCII hyphen
# ─────────────────────────────────────────────
def norm(s):
    return s.replace("\u2212", "-").replace("\u2013", "-")


# ─────────────────────────────────────────────
# HEADER EXTRACTION
# ─────────────────────────────────────────────
def extract_lab_info(text):
    lab     = re.search(r"Lab[:\s]+(\d+)",                              text, re.IGNORECASE)
    cycle   = re.search(r"Cycle\s+(\d+)",                               text, re.IGNORECASE)
    sample  = re.search(r"Sample\s+No[:\s]+(\d+)",                      text, re.IGNORECASE)
    sdate   = re.search(r"Sample\s+Date[:\s]+([\d]+\s+\w+\s+[\d]+)", text, re.IGNORECASE)
    lot     = re.search(r"Lot\s+No[:\s]+(\d+)",                         text, re.IGNORECASE)
    # Program: from "Lab XXXXX <Program Name> Cycle N" header line
    program = re.search(r"Lab\s+\d+\s+(.+?Program)\s+Cycle",           text, re.IGNORECASE)
    # Fallback: standalone title "Cardiac Markers Program (BC39)"
    if not program:
        program = re.search(r"^(.+?Program)(?:\s*\([^)]+\))?(?:\s+Cycle|\s*$)",
                            text, re.IGNORECASE | re.MULTILINE)
    return {
        "Lab":         lab.group(1)              if lab     else None,
        "Cycle":       cycle.group(1)            if cycle   else None,
        "Sample":      sample.group(1)           if sample  else None,
        "Sample Date": sdate.group(1)            if sdate   else None,
        "Lot No":      lot.group(1)              if lot     else None,
        "Program":     program.group(1).strip()  if program else None,
    }


# ─────────────────────────────────────────────
# SUMMARY PAGE PARSER
#
# Uses a lazy analyte match (.+?) so names with commas, slashes,
# and parentheses all work:
#   "Bilirubin, Direct"   "ALT (ALAT/GPT)"   "AST/GOT"
#   "Cholesterol, LDL"    "Bilirubin, Indirect/BU"   "eGFR"
# Unit pattern explicitly handles mL/min/1.73m2
# ─────────────────────────────────────────────
SUMMARY_ROW = re.compile(
    r"^(.+?)"
    r"\s+"
    r"((?:mL/min/1\.73m2"           # eGFR compound unit
    r"|fL\s+\(cubic\s+µm\)"       # hematology fL (cubic µm)
    r"|pg/cell|K/µL|M/µL"            # hematology cell-count units
    r"|[a-zA-Z%µ][a-zA-Z0-9%µ]*(?:/[a-zA-Z0-9\.µ]+)*))"
    r"\s+"
    r"([\d\.]+)"
    r"\s+"
    r"([\d\.]+)"
    r"\s+"
    r"([\-\d\.]+)"
    r"\s+"
    r"([\-\d\.]+)"
    r"\s+Peer",
    re.IGNORECASE
)

def parse_summary_page(text, lab_info):
    text = norm(text)
    records = []
    for line in text.split("\n"):
        m = SUMMARY_ROW.match(line.strip())
        if m:
            analyte, unit, result, mean, zscore, rmz = m.groups()
            records.append({
                "Lab":         lab_info["Lab"],
                "Cycle":       lab_info["Cycle"],
                "Sample":      lab_info["Sample"],
                "Sample Date": lab_info.get("Sample Date"),
                "Lot No":      lab_info.get("Lot No"),
                "Program":     lab_info.get("Program"),
                "Analyte":     re.sub(r"\.{2,}$|…$", "", analyte.strip()).strip(),
                "Unit":        unit.strip(),
                "Result":      float(result),
                "Peer Mean":   float(mean),
                "Z-score":     float(zscore),
                "RMZ":         float(rmz),
            })
    return records


# ─────────────────────────────────────────────
# INDIVIDUAL ANALYTE PAGE PARSER (fallback)
# ─────────────────────────────────────────────
PEER_ROW = re.compile(
    r"Your\s+Peer\s+"
    r"(\d+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\-\d\.]+)\s+([\-\d\.]+)",
    re.IGNORECASE
)

RESULT_LINE = re.compile(
    r"^([\d\.]+)\s+"
    r"(fL\s+\(cubic\s+µm\)"      # hematology fL (cubic µm)
    r"|mL/min/1\.73m2"
    r"|pg/mL|pg/dL|pg/cell"
    r"|mg/L|mg/dL|ng/L|ng/mL|ng/dL"
    r"|µg/L|µg/mL|µg/dL|g/L|g/dL"
    r"|K/µL|M/µL"
    r"|U/L|IU/L|IU/mL|mIU/L|mIU/mL|µIU/mL|µIU/L"
    r"|mmol/L|µmol/L|nmol/L|pmol/L|%|ratio)",
    re.IGNORECASE
)

SKIP_TITLES = {
    "configuration report", "sample summary report",
    "data on file report", "exceptions", "sample comments",
}

def parse_analyte_page(text):
    text = norm(text)
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    analyte = None
    for line in lines[:3]:
        if "report" in line.lower() and len(line) < 80:
            candidate = re.sub(r"\s*report\s*$", "", line, flags=re.IGNORECASE).strip()
            if candidate.lower() not in SKIP_TITLES and len(candidate) > 1:
                analyte = candidate
            break
    if not analyte:
        return None

    result = None
    for line in lines:
        m = RESULT_LINE.match(line)
        if m:
            result = float(m.group(1))
            break

    peer_mean = zscore = rmz = None
    pm = PEER_ROW.search(text)
    if pm:
        peer_mean = float(pm.group(2))
        zscore    = float(pm.group(6))
        rmz       = float(pm.group(7))

    if result is None:
        return None

    return {
        "analyte":   analyte,
        "result":    result,
        "peer_mean": peer_mean,
        "zscore":    zscore,
        "rmz":       rmz,
    }


# ─────────────────────────────────────────────
# MAIN EXTRACTION
# ─────────────────────────────────────────────
def extract_all(file):
    base_records = {}
    lab_info = {"Lab": None, "Cycle": None, "Sample": None, "Sample Date": None, "Lot No": None, "Program": None}

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            info = extract_lab_info(text)
            for k in ("Lab", "Cycle", "Sample", "Sample Date", "Lot No", "Program"):
                if info[k] and not lab_info[k]:
                    lab_info[k] = info[k]

            if "sample summary report" in text.lower():
                for rec in parse_summary_page(text, lab_info):
                    rec["Analyte"] = norm(rec["Analyte"])
                    base_records[rec["Analyte"].lower()] = rec

            detail = parse_analyte_page(text)
            if detail:
                detail["analyte"] = norm(detail["analyte"])
                key = detail["analyte"].lower()
                # Also check if any summary key is a prefix of this detail name
                # e.g. "microalbumin/urine alb" matches "microalbumin/urine albumin"
                if key not in base_records:
                    for existing_key in list(base_records.keys()):
                        if key.startswith(existing_key) or existing_key.startswith(key):
                            # Upgrade the truncated name to the full name from the detail page
                            rec = base_records.pop(existing_key)
                            rec["Analyte"] = detail["analyte"]
                            base_records[key] = rec
                            break
                if key not in base_records:
                    base_records[key] = {
                        "Lab":         lab_info["Lab"],
                        "Cycle":       lab_info["Cycle"],
                        "Sample":      lab_info["Sample"],
                        "Sample Date": lab_info.get("Sample Date"),
                        "Lot No":      lab_info.get("Lot No"),
                        "Program":     lab_info.get("Program"),
                        "Analyte":     detail["analyte"],
                        "Unit":        None,
                        "Result":      detail["result"],
                        "Peer Mean":   detail["peer_mean"],
                        "Z-score":     detail["zscore"],
                        "RMZ":         detail["rmz"],
                    }

    for rec in base_records.values():
        for k in ("Lab", "Cycle", "Sample", "Sample Date", "Lot No", "Program"):
            if not rec.get(k):
                rec[k] = lab_info.get(k)

    if not base_records:
        return pd.DataFrame()

    cols = ["Lab", "Cycle", "Sample", "Sample Date", "Lot No", "Program",
            "Analyte", "Unit", "Result", "Peer Mean", "Z-score", "RMZ"]
    return pd.DataFrame(list(base_records.values()))[cols]


# ─────────────────────────────────────────────
# GOOGLE SHEETS UPLOAD
# ─────────────────────────────────────────────
def _cast(v):
    """
    Convert a value to the type gspread should store:
    - None / NaN          → empty string (no backtick)
    - int / float         → number as-is
    - purely numeric str  → int  (Lab, Cycle, Sample, Lot No — prevents backtick)
    - everything else     → str  (Analyte, Unit, Program, Sample Date)
    gspread only prepends a backtick when it sends a str that looks numeric,
    so casting those to int eliminates the issue completely.
    """
    if v is None:
        return ""
    if isinstance(v, float) and v != v:   # NaN
        return ""
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if s.lstrip("-").isdigit():            # purely integer string → int
        return int(s)
    # Reformat "DD-Mon-YY" or "DD Mon YY" dates so no Sheets locale
    # can interpret them as dates → store as plain text, never a date cell
    import re as _re
    dm = _re.match(
        r"^(\d{1,2})[\s\-](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\s\-](\d{2,4})$",
        s, _re.IGNORECASE)
    if dm:
        d, mon, y = dm.group(1), dm.group(2).capitalize(), dm.group(3)
        # Format as DD-MM-YY (numeric month) so Sheets stores as plain text
        month_map = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04","May":"05",
                     "Jun":"06","Jul":"07","Aug":"08","Sep":"09","Oct":"10",
                     "Nov":"11","Dec":"12"}
        mm = month_map.get(mon, mon)
        yy = y[-2:]  # always 2-digit year
        return f"{int(d):02d}-{mon}-{yy}"
    return s                               # text stays as text — no backtick risk


def _get_or_create_sheet(spreadsheet, lab_id):
    """Return the worksheet for this lab, creating it with a header if new."""
    tab_name = str(lab_id)
    try:
        return spreadsheet.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=tab_name, rows=1000, cols=20)
        return ws


def upload_to_gsheets(df):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds  = ServiceAccountCredentials.from_json_keyfile_dict(
        st.secrets["gcp_service_account"], scope
    )
    client       = gspread.authorize(creds)
    spreadsheet  = client.open("EQAS Master Dashboard")
    headers      = df.columns.tolist()
    total_added  = 0

    # Process each lab separately into its own tab
    for lab_id, lab_df in df.groupby("Lab"):
        ws = _get_or_create_sheet(spreadsheet, lab_id)

        existing = ws.get_all_records()

        # Write header if sheet is empty
        if not existing:
            ws.append_rows([headers], value_input_option="RAW")

        existing_set = set(
            (str(r.get("Lab", "")), str(r.get("Cycle", "")), str(r.get("Sample", "")),
             str(r.get("Sample Date", "")), str(r.get("Lot No", "")), r.get("Analyte", ""))
            for r in existing
        )

        new_rows = []
        for _, row in lab_df.iterrows():
            key = (str(row["Lab"]), str(row["Cycle"]), str(row["Sample"]),
                   str(row["Sample Date"]), str(row["Lot No"]), row["Analyte"])
            if key not in existing_set:
                new_rows.append([_cast(v) for v in row.tolist()])

        if new_rows:
            ws.append_rows(new_rows, value_input_option="RAW")
            total_added += len(new_rows)

    return total_added


# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────
if uploaded_files:
    all_data = []

    for file in uploaded_files:
        with st.spinner(f"Parsing {file.name} …"):
            df = extract_all(file)

        if df.empty:
            st.warning(
                f"⚠️ No data extracted from **{file.name}**. "
                "Use the debug expander below to inspect the raw text."
            )
        else:
            all_data.append(df)

        with st.expander(f"🔍 Debug: raw PDF text — {file.name}"):
            with pdfplumber.open(file) as pdf:
                for i, page in enumerate(pdf.pages):
                    st.markdown(f"**Page {i+1}**")
                    st.text(page.extract_text() or "(no text extracted)")

    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)

        st.subheader("Extracted Data")
        st.dataframe(final_df, use_container_width=True)
        st.caption(
            f"✅ {len(final_df)} analyte row(s) extracted from {len(uploaded_files)} file(s)"
        )

        from io import BytesIO
        col1, col2 = st.columns(2)

        with col1:
            output = BytesIO()
            final_df.to_excel(output, index=False, engine="openpyxl")
            output.seek(0)
            st.download_button(
                "⬇️ Download Excel",
                data=output,
                file_name="EQAS_Extract.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with col2:
            if st.button("☁️ Upload to Google Sheets"):
                try:
                    added = upload_to_gsheets(final_df)
                    st.success(f"Uploaded — {added} new row(s) added (duplicates skipped)")
                except Exception as e:
                    st.error(f"Upload failed: {e}")
