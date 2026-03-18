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
    lab    = re.search(r"Lab[:\s]+(\d+)",                            text, re.IGNORECASE)
    cycle  = re.search(r"Cycle\s+(\d+)",                             text, re.IGNORECASE)
    sample = re.search(r"Sample\s+No[:\s]+(\d+)",                    text, re.IGNORECASE)
    sdate  = re.search(r"Sample\s+Date[:\s]+([\d]+\s+\w+\s+[\d]+)", text, re.IGNORECASE)
    return {
        "Lab":         lab.group(1)    if lab    else None,
        "Cycle":       cycle.group(1)  if cycle  else None,
        "Sample":      sample.group(1) if sample else None,
        "Sample Date": sdate.group(1)  if sdate  else None,
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
    r"((?:mL/min/1\.73m2|[a-zA-Z%µ][a-zA-Z0-9%µ]*(?:/[a-zA-Z0-9\.µ]+)*))"
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
                "Analyte":     analyte.strip(),
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
    r"(mL/min/1\.73m2"
    r"|pg/mL|pg/dL|mg/L|mg/dL|ng/L|ng/mL|ng/dL"
    r"|µg/L|µg/mL|µg/dL|g/L|g/dL"
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
    lab_info = {"Lab": None, "Cycle": None, "Sample": None, "Sample Date": None}

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            info = extract_lab_info(text)
            for k in ("Lab", "Cycle", "Sample", "Sample Date"):
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
                if key not in base_records:
                    base_records[key] = {
                        "Lab":         lab_info["Lab"],
                        "Cycle":       lab_info["Cycle"],
                        "Sample":      lab_info["Sample"],
                        "Sample Date": lab_info.get("Sample Date"),
                        "Analyte":     detail["analyte"],
                        "Unit":        None,
                        "Result":      detail["result"],
                        "Peer Mean":   detail["peer_mean"],
                        "Z-score":     detail["zscore"],
                        "RMZ":         detail["rmz"],
                    }

    for rec in base_records.values():
        for k in ("Lab", "Cycle", "Sample", "Sample Date"):
            if not rec.get(k):
                rec[k] = lab_info.get(k)

    if not base_records:
        return pd.DataFrame()

    cols = ["Lab", "Cycle", "Sample", "Sample Date", "Analyte", "Unit",
            "Result", "Peer Mean", "Z-score", "RMZ"]
    return pd.DataFrame(list(base_records.values()))[cols]


# ─────────────────────────────────────────────
# GOOGLE SHEETS UPLOAD
# ─────────────────────────────────────────────
def _cast(v):
    if v is None:
        return ""
    if isinstance(v, float) and v != v:
        return ""
    if isinstance(v, (int, float)):
        return v
    try:
        return int(v)
    except (ValueError, TypeError):
        pass
    return str(v)


def upload_to_gsheets(df):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds  = ServiceAccountCredentials.from_json_keyfile_dict(
        st.secrets["gcp_service_account"], scope
    )
    client = gspread.authorize(creds)
    sheet  = client.open("EQAS Master Dashboard").sheet1

    existing     = sheet.get_all_records()
    existing_set = set(
        (str(r.get("Lab", "")), str(r.get("Cycle", "")), str(r.get("Sample", "")),
         str(r.get("Sample Date", "")), r.get("Analyte", ""))
        for r in existing
    )

    if not existing:
        sheet.append_row(df.columns.tolist())

    new_rows = []
    for _, row in df.iterrows():
        key = (str(row["Lab"]), str(row["Cycle"]), str(row["Sample"]),
               str(row["Sample Date"]), row["Analyte"])
        if key not in existing_set:
            new_rows.append([_cast(v) for v in row.tolist()])

    for row in new_rows:
        sheet.append_row(row)

    return len(new_rows)


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
