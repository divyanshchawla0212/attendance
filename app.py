import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import re

st.set_page_config(page_title="KollegeApply Attendance Summary Generator")
st.title("📋 KollegeApply Attendance Summary Generator")
st.caption("Upload Attendance Excel File (.xlsx or .xls)")

uploaded_file = st.file_uploader("Upload file", type=["xlsx", "xls"])

def try_parse_date(date_str):
    for fmt in ("%d-%b-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(date_str.strip(), fmt).strftime("%Y-%m-%d")
        except:
            continue
    return datetime.today().strftime("%Y-%m-%d")

def detect_log_format(buffer, engine):
    try:
        df = pd.read_excel(buffer, engine=engine)
        if "Original record" in df.columns:
            return df
    except:
        return None
    return None

def detect_tabular_format(buffer, engine):
    try:
        df = pd.read_excel(buffer, header=4, engine=engine)
        tabular_cols = {"E. Code", "Name", "InTime", "OutTime", "Status"}
        if tabular_cols.issubset(set(df.columns)):
            return df
    except:
        return None
    return None

if uploaded_file:
    try:
        file_ext = uploaded_file.name.split(".")[-1].lower()
        engine = "xlrd" if file_ext == "xls" else "openpyxl"
        buffer = BytesIO(uploaded_file.read())
        buffer.seek(0)

        summary = None
        full_date = datetime.today().strftime("%Y-%m-%d")

        # CASE 1: Raw log format
        df_log = detect_log_format(buffer, engine)
        if df_log is not None:
            st.success("Detected Log Format (Original record)")
            df_log = df_log.dropna(subset=['Original record']).reset_index(drop=True)

            date_line = str(df_log.loc[0, 'Original record'])
            date_match = re.search(r'Date:(\d{4}-\d{1,2}-\d{1,2})', date_line)
            full_date = datetime.strptime(date_match.group(1), "%Y-%m-%d").strftime("%Y-%m-%d") if date_match else full_date

            records = []
            for i in range(1, len(df_log) - 2, 3):
                try:
                    person_info = df_log.loc[i, 'Original record']
                    time_entries = df_log.loc[i + 2, 'Original record']

                    name = re.search(r'Name:(.*?)Dept', person_info)
                    dept = re.search(r'Dept\.:([^\s]+)', person_info)

                    name = name.group(1).strip() if name else "Unknown"
                    dept = dept.group(1).strip() if dept else "Unknown"

                    times = [t.strip() for t in time_entries.strip().split('\n') if t.strip()]
                    times_dt = [datetime.strptime(f"{full_date} {t}", "%Y-%m-%d %H:%M") for t in times]

                    in_time = min(times_dt).strftime("%H:%M")
                    out_time = max(times_dt).strftime("%H:%M")

                    records.append({
                        "Name": name,
                        "Department": dept,
                        "Date": full_date,
                        "In Time": in_time,
                        "Out Time": out_time
                    })
                except:
                    continue

            summary = pd.DataFrame(records)

        else:
            # CASE 2: Tabular sheet format
            buffer.seek(0)
            df_tab = detect_tabular_format(buffer, engine)
            if df_tab is not None:
                st.success("Detected Tabular Daily Attendance Format")

                buffer.seek(0)
                df_header = pd.read_excel(buffer, header=None, engine=engine)
                raw_date = df_header.iloc[1, 1]
                full_date = try_parse_date(str(raw_date))

                summary = df_tab[["E. Code", "Name", "InTime", "OutTime", "Status"]].copy()
                summary.rename(columns={"E. Code": "Emp Code"}, inplace=True)
                summary["Date"] = full_date
                summary = summary[["Emp Code", "Name", "Date", "InTime", "OutTime", "Status"]]

        if summary is not None and not summary.empty:
            st.dataframe(summary)
            csv_data = summary.to_csv(index=False).encode("utf-8")
            st.download_button("📥 Download Summary CSV", csv_data, file_name=f"summary_{full_date}.csv", mime="text/csv")
        else:
            st.warning("Unable to detect supported format. Please check the file structure.")

    except Exception as e:
        st.error("Something went wrong. Try with a valid Excel file.")
        st.exception(e)
