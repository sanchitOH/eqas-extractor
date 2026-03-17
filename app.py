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
# HEADER EXTRACTION  (Lab / Cycle / Sample No)
# Works on any page header like:
#   "Lab 328024 Cardiac Markers Program Cycle 13"
#   "Lab: 328024"   "Sample No: 11"
# ─────────────────────────────────────────────
def extract_lab_info(text):
    lab    = re.search(r"Lab[:\s]+(\d+)",                    text, re.IGNORECASE)
    cycle  = re.search(r"Cycle\s+(\d+)",                     text, re.IGNORECASE)
    sample = re.search(r"Sample\s+No[:\s]+(\d+)",            text, re.IGNORECASE)

    return {
        "Lab":    lab.group(1)    if lab    else None,
        "Cycle":  cycle.group(1)  if cycle  else None,
        "Sample": sample.group(1) if sample else None,
    }


# ─────────────────────────────────────────────
# STRATEGY A — parse the "Sample Summary Report"
# (page 3 in this PDF).  One compact table has
# every analyte with Result / Mean / RMZ.
#
# Raw text pattern per analyte row (after ✔/✘):
#   NT−proBNP pg/mL 2571 2328 1.56 1.56 Peer
#   hs−CRP    mg/L  11.378 10.9 0.94 0.19 Peer
#   Troponin−I ng/L 18380 17522 0.89 −0.81 Peer
#
# Columns: Analyte  Unit  Result  Mean  Z-score  RMZ  Comparator
# ─────────────────────────────────────────────
SUMMARY_ROW = re.compile(
    r"^([\w\-−]+(?:\s[\w\-−]+)*)"          # analyte name (may contain spaces/hyphens)
    r"\s+"
    r"([a-zA-Z/%µ]+(?:/[a-zA-Z]+)?)"       # unit  e.g. pg/mL  mg/L  ng/L  U/L
    r"\s+"
    r"([\d\.]+)"                            # Result
    r"\s+"
    r"([\d\.]+)"                            # Mean
    r"\s+"
    r"([\-\d\.]+)"                          # Z-score
    r"\s+"
    r"([\-\d\.]+)"                          # RMZ
    r"\s+Peer",
    re.IGNORECASE
)

def parse_summary_page(text, lab_info):
    """
    Parse the compact summary table present on the 'Sample Summary Report' page.
    Returns a list of dicts, one per analyte.
    """
    # Normalise unicode minus (U+2212) to ASCII hyphen so numeric regexes work
    text = text.replace("\u2212", "-").replace("\u2013", "-")
    records = []
    for line in text.split("\n"):
        line = line.strip()
        m = SUMMARY_ROW.match(line)
        if m:
            analyte, unit, result, mean, zscore, rmz = m.groups()
            # For Peer SD we don't have it on the summary page; fill from detail pages
            records.append({
                "Lab":       lab_info["Lab"],
                "Cycle":     lab_info["Cycle"],
                "Sample":    lab_info["Sample"],
                "Analyte":   analyte.strip(),
                "Unit":      unit.strip(),
                "Result":    float(result),
                "Peer Mean": float(mean),
                "Peer SD":   None,          # filled below from individual report pages
                "RMZ":       float(rmz),
            })
    return records


# ─────────────────────────────────────────────
# STRATEGY B — parse individual analyte pages
# ("NT−proBNP Report", "hs−CRP Report", …)
#
# Key patterns found in raw text:
#
#  Page title:  "NT−proBNP Report"  (first non-empty line)
#
#  Your Result: "2571 pg/mL"  appears right after "Your Result"
#
#  Your Peer row:
#    "Your Peer 112 2328 156 6.69 36.8 1.56 1.56 10.4"
#    columns: N  Mean  SD  CV  U  Z-score  RMZ  %Dev
# ─────────────────────────────────────────────

# Matches: "Your Peer  112  2328  156  6.69  36.8  1.56  1.56  10.4"
PEER_ROW = re.compile(
    r"Your\s+Peer\s+"
    r"(\d+)\s+"           # N
    r"([\d\.]+)\s+"       # Mean
    r"([\d\.]+)\s+"       # SD   ← this is what we want
    r"([\d\.]+)\s+"       # CV
    r"([\d\.]+)\s+"       # U¹
    r"([\-\d\.]+)\s+"     # Z-score
    r"([\-\d\.]+)",       # RMZ
    re.IGNORECASE
)

# Matches the result line: "2571 pg/mL Your Method ..."  OR  "11.378 mg/L Your Method ..."
# The value always appears at the START of a line, followed by the unit, then more text.
# Pattern covers common units; extend the list for other EQAS programs.
RESULT_LINE = re.compile(
    r"^([\d\.]+)\s+"
    r"(pg/mL|mg/L|ng/L|µg/L|g/L|U/L|IU/L|mIU/L|mmol/L|µmol/L|nmol/L|pmol/L|%|ratio)",
    re.IGNORECASE
)

def parse_analyte_page(text):
    """
    Parse a single-analyte detail page.
    Returns dict with keys: analyte, result, peer_mean, peer_sd, rmz
    or None if the page is not a report page.

    Layout (confirmed from real PDF):
      Line 0 : "NT−proBNP Report"           ← analyte name
      Line 8 : "Your Result … N Mean SD …"  ← table header (contains (cid:) artifacts)
      Line 10: "(cid:2)(cid:3)"             ← arrow artifact — SKIP
      Line 11: "2571 pg/mL Your Method …"   ← YOUR RESULT IS HERE (starts the line)
      Line 13: "Your Peer 112 2328 156 …"   ← peer stats row
    """
    # Normalise unicode minus (U+2212) to ASCII hyphen so numeric regexes work
    text = text.replace("\u2212", "-").replace("\u2013", "-")
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    # ── Analyte name ──────────────────────────────────────────────────────────
    # First non-empty line is always "<AnalyteName> Report"
    analyte = None
    skip_titles = {"configuration report", "sample summary report",
                   "data on file report", "exceptions"}
    for line in lines[:3]:
        if "report" in line.lower() and len(line) < 60:
            candidate = re.sub(r'\s*report\s*$', '', line, flags=re.IGNORECASE).strip()
            candidate = candidate.replace("−", "-")
            if candidate.lower() not in skip_titles and len(candidate) > 1:
                analyte = candidate
            break
    if not analyte:
        return None

    # ── Your Result ───────────────────────────────────────────────────────────
    # The value sits on a line that STARTS with the numeric result + unit.
    # e.g. "2571 pg/mL Your Method 142 2390 …"
    result = None
    for line in lines:
        m = RESULT_LINE.match(line)
        if m:
            result = float(m.group(1))
            break

    # ── Your Peer row ─────────────────────────────────────────────────────────
    # "Your Peer  N  Mean  SD  CV  U¹  Z-score  RMZ  %Dev"
    peer_mean = peer_sd = rmz = None
    peer_m = PEER_ROW.search(text)
    if peer_m:
        peer_mean = float(peer_m.group(2))
        peer_sd   = float(peer_m.group(3))
        rmz       = float(peer_m.group(7))

    if result is None:
        return None

    return {
        "analyte":   analyte,
        "result":    result,
        "peer_mean": peer_mean,
        "peer_sd":   peer_sd,
        "rmz":       rmz,
    }


# ─────────────────────────────────────────────
# MAIN EXTRACTION — two-pass approach
#   Pass 1: find the summary page → creates base records dict keyed by analyte
#   Pass 2: scan individual report pages → fills in Peer SD (and cross-checks)
# ─────────────────────────────────────────────
def extract_all(file):
    base_records = {}   # analyte → record dict
    lab_info = {"Lab": None, "Cycle": None, "Sample": None}

    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text:
                continue

            # Update lab_info whenever we find the values
            info = extract_lab_info(text)
            for k in ("Lab", "Cycle", "Sample"):
                if info[k] and not lab_info[k]:
                    lab_info[k] = info[k]

            # ── Summary page ───────────────────────────────────────────────
            if "sample summary report" in text.lower():
                for rec in parse_summary_page(text, lab_info):
                    # Normalise typographic dash so keys match detail pages
                    rec["Analyte"] = rec["Analyte"].replace("\u2212", "-").replace("\u2013", "-")
                    key = rec["Analyte"].lower()
                    base_records[key] = rec

            # ── Individual analyte report page ────────────────────────────
            detail = parse_analyte_page(text)
            if detail:
                # Normalise typographic dash so keys match summary page
                detail["analyte"] = detail["analyte"].replace("\u2212", "-").replace("\u2013", "-")
                key = detail["analyte"].lower()
                if key in base_records:
                    # Enrich existing record with Peer SD (and Unit if missing)
                    base_records[key]["Peer SD"] = detail["peer_sd"]
                else:
                    # No summary page found yet — create record from detail
                    base_records[key] = {
                        "Lab":       lab_info["Lab"],
                        "Cycle":     lab_info["Cycle"],
                        "Sample":    lab_info["Sample"],
                        "Analyte":   detail["analyte"],
                        "Unit":      None,
                        "Result":    detail["result"],
                        "Peer Mean": detail["peer_mean"],
                        "Peer SD":   detail["peer_sd"],
                        "RMZ":       detail["rmz"],
                    }

    # Back-fill lab info in case summary was hit before the header
    for rec in base_records.values():
        for k in ("Lab", "Cycle", "Sample"):
            if not rec.get(k):
                rec[k] = lab_info.get(k)

    if not base_records:
        return pd.DataFrame()

    cols = ["Lab", "Cycle", "Sample", "Analyte", "Unit",
            "Result", "Peer Mean", "Peer SD", "RMZ"]
    return pd.DataFrame(list(base_records.values()))[cols]


# ─────────────────────────────────────────────
# GOOGLE SHEETS UPLOAD
# ─────────────────────────────────────────────
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
        (str(r["Lab"]), str(r["Cycle"]), str(r["Sample"]), r["Analyte"])
        for r in existing
    )

    if not existing:
        sheet.append_row(df.columns.tolist())

    new_rows = []
    for _, row in df.iterrows():
        key = (str(row["Lab"]), str(row["Cycle"]), str(row["Sample"]), row["Analyte"])
        if key not in existing_set:
            new_rows.append([str(v) if v is not None else "" for v in row.tolist()])

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

        # Debug expander — shows raw text for every page so you can tune regexes
        with st.expander(f"🔍 Debug: raw PDF text — {file.name}"):
            with pdfplumber.open(file) as pdf:
                for i, page in enumerate(pdf.pages):
                    st.markdown(f"**Page {i+1}**")
                    st.text(page.extract_text() or "(no text extracted)")

    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)

        st.subheader("Extracted Data")
        st.dataframe(final_df, use_container_width=True)
        st.caption(f"✅ {len(final_df)} analyte row(s) extracted from {len(uploaded_files)} file(s)")

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
