import streamlit as st
import pdfplumber
import pandas as pd
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="EQAS Extractor", layout="wide")
st.title("EQAS Auto Extractor")

uploaded_files = st.file_uploader("Upload EQAS PDF(s)", type=["pdf"], accept_multiple_files=True)


def extract_lab_info(text):
    lab = re.search(r"Lab[:\s]+(\d+)", text)
    cycle = re.search(r"Cycle\s+(\d+)", text)
    sample = re.search(r"Sample No[:\s]+(\d+)", text)

    return {
        "Lab": lab.group(1) if lab else None,
        "Cycle": cycle.group(1) if cycle else None,
        "Sample": sample.group(1) if sample else None
    }


def clean_analyte_name(name):
    name = name.replace("−", "-")
    return name.strip()


def extract_metrics(text):
    records = []

    # Split into sections using "Your Result"
    sections = re.split(r"Your Result", text, flags=re.IGNORECASE)

    for section in sections[1:]:  # skip first junk part
        try:
            # -------------------------
            # 1. Extract RESULT
            # -------------------------
            result_match = re.search(r"([\d]+\.\d+|[\d]+)", section)
            result = float(result_match.group(1)) if result_match else None

            # -------------------------
            # 2. Extract PEER DATA
            # -------------------------
            peer_match = re.search(
                r"Your Peer\s+\d+\s+([\d\.]+)\s+([\d\.]+).*?([\-\d\.]+)\s+([\-\d\.]+)",
                section,
                re.DOTALL | re.IGNORECASE
            )

            peer_mean = float(peer_match.group(1)) if peer_match else None
            peer_sd = float(peer_match.group(2)) if peer_match else None
            rmz = float(peer_match.group(4)) if peer_match else None

            # -------------------------
            # 3. Extract ANALYTE (look BEFORE this section)
            # -------------------------
            prev_text = sections[sections.index(section)-1]

            lines = prev_text.strip().split("\n")
            lines = [l.strip() for l in lines if l.strip()]

            analyte = None
            for line in reversed(lines[-10:]):  # look last few lines
                if (
                    len(line) < 50
                    and not re.search(r"\d", line)
                    and "report" not in line.lower()
                    and "configuration" not in line.lower()
                ):
                    analyte = line
                    break

            # -------------------------
            # 4. Save record
            # -------------------------
            if analyte and result:
                records.append({
                    "Analyte": analyte,
                    "Result": result,
                    "Peer Mean": peer_mean,
                    "Peer SD": peer_sd,
                    "RMZ": rmz
                })

        except:
            continue

    return records


def extract_all(pdf):
    final = []

    with pdfplumber.open(pdf) as pdf_file:
        for page in pdf_file.pages:
            text = page.extract_text()
            if not text:
                continue

            info = extract_lab_info(text)
            metrics = extract_metrics(text)

            for m in metrics:
                final.append({
                    "Lab": info["Lab"],
                    "Cycle": info["Cycle"],
                    "Sample": info["Sample"],
                    **m
                })

    return pd.DataFrame(final)


def upload_to_gsheets(df):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        st.secrets["gcp_service_account"], scope
    )

    client = gspread.authorize(creds)
    sheet = client.open("EQAS Master Dashboard").sheet1

    existing = sheet.get_all_records()
    existing_set = set(
        (str(r["Lab"]), str(r["Cycle"]), str(r["Sample"]), r["Analyte"])
        for r in existing
    )

    new_rows = []
    for _, row in df.iterrows():
        key = (str(row["Lab"]), str(row["Cycle"]), str(row["Sample"]), row["Analyte"])
        if key not in existing_set:
            new_rows.append(row.tolist())

    if not existing:
        sheet.append_row(df.columns.tolist())

    for row in new_rows:
        sheet.append_row(row)


if uploaded_files:
    all_data = []

    for file in uploaded_files:
        df = extract_all(file)
        all_data.append(df)

    final_df = pd.concat(all_data, ignore_index=True)

    st.subheader("Extracted Data")
    st.dataframe(final_df)

    from io import BytesIO

    col1, col2 = st.columns(2)

    with col1:
        output = BytesIO()
        final_df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            "Download Excel",
            data=output,
            file_name="EQAS_Extract.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col2:
        if st.button("Upload to Google Sheets"):
            upload_to_gsheets(final_df)
            st.success("Uploaded (duplicates avoided)")
