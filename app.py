import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import zipfile
import re

st.set_page_config(page_title="KollegeApply Attendance Summary Generator")
st.title("📋 KollegeApply Attendance Summary Generator")
st.caption("Upload Attendance Excel File (.xlsx or .xls)")

uploaded_file = st.file_uploader("Upload file", type=["xlsx", "xls"])

def parse_sno_format(buffer):
    df_all = pd.read_excel(buffer, sheet_name=0, header=None, engine="openpyxl")
    try:
        attendance_date = df_all.iloc[1, 4]
        if isinstance(attendance_date, datetime):
            full_date = attendance_date.strftime("%Y-%m-%d")
        elif isinstance(attendance_date, str):
            full_date = datetime.strptime(attendance_date.strip(), "%d-%b-%Y").strftime("%Y-%m-%d")
        else:
            full_date = datetime.today().strftime("%Y-%m-%d")
    except:
        full_date = datetime.today().strftime("%Y-%m-%d")

    df = pd.read_excel(buffer, sheet_name=0, header=4, engine="openpyxl")
    if not {"E. Code", "Name", "InTime", "OutTime", "Status"}.intersection(df.columns):
        return None, None

    df_summary = df[["E. Code", "Name", "InTime", "OutTime", "Status"]].copy()
    df_summary.rename(columns={"E. Code": "Emp Code"}, inplace=True)
    df_summary["Date"] = full_date
    df_summary = df_summary[["Emp Code", "Name", "Date", "InTime", "OutTime", "Status"]]
    return df_summary, full_date

def parse_name_status_format(buffer):
    df_excel = pd.read_excel(buffer, sheet_name=0, header=4, engine="openpyxl")
    if {"Name", "In Time", "OutTime", "Status"}.issubset(df_excel.columns):
        df_raw = pd.read_excel(buffer, sheet_name=0, header=None, engine="openpyxl")
        raw_date = df_raw.iloc[1, 1]
        if isinstance(raw_date, datetime):
            full_date = raw_date.strftime("%Y-%m-%d")
        elif isinstance(raw_date, str):
            full_date = datetime.strptime(raw_date.strip(), "%d-%b-%Y").strftime("%Y-%m-%d")
        else:
            full_date = datetime.today().strftime("%Y-%m-%d")

        df_summary = df_excel[["E. Code", "Name", "In Time", "OutTime", "Status"]].copy()
        df_summary.rename(columns={"E. Code": "Emp Code"}, inplace=True)
        df_summary["Date"] = full_date
        df_summary = df_summary[["Emp Code", "Name", "Date", "In Time", "OutTime", "Status"]]
        return df_summary, full_date
    return None, None

def parse_raw_format(buffer):
    df_clean = pd.read_excel(buffer, engine="openpyxl")
    if 'Original record' not in df_clean.columns:
        return None, None

    df_clean = df_clean.dropna(subset=['Original record']).reset_index(drop=True)
    date_line = str(df_clean.loc[0, 'Original record'])
    date_match = re.search(r'Date:(\d{4}-\d{1,2}-\d{1,2})', date_line)
    full_date = datetime.strptime(date_match.group(1), "%Y-%m-%d").strftime("%Y-%m-%d") if date_match else datetime.today().strftime("%Y-%m-%d")

    records = []
    for i in range(1, len(df_clean) - 2, 3):
        try:
            person_info = df_clean.loc[i, 'Original record']
            time_entries = df_clean.loc[i + 2, 'Original record']

            name_match = re.search(r'Name:(.*?)Dept', person_info)
            name = name_match.group(1).strip() if name_match else "Unknown"

            dept_match = re.search(r'Dept\.:([^\s]+)', person_info)
            dept = dept_match.group(1).strip() if dept_match else "Unknown"

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

        except Exception:
            continue

    df_summary = pd.DataFrame(records)
    return df_summary, full_date


if uploaded_file:
    try:
        buffer = BytesIO(uploaded_file.read())
        if not zipfile.is_zipfile(buffer):
            st.error("Uploaded file is not a valid Excel .xlsx file. Please re-save it in Excel.")
            st.stop()
        buffer.seek(0)

        df_summary, full_date = parse_sno_format(buffer)
        if df_summary is None:
            buffer.seek(0)
            df_summary, full_date = parse_name_status_format(buffer)
        if df_summary is None:
            buffer.seek(0)
            df_summary, full_date = parse_raw_format(buffer)

        if df_summary is None:
            st.error("❌ Unable to detect supported format. Please check the file structure.")
        else:
            st.success("✅ File processed successfully!")
            st.dataframe(df_summary)

            csv_buffer = BytesIO()
            df_summary.to_csv(csv_buffer, index=False)
            csv_buffer.seek(0)

            st.download_button(
                label=f"📥 Download CSV (summary_{full_date}.csv)",
                data=csv_buffer,
                file_name=f"df_summary_{full_date}.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"Error: {e}")
