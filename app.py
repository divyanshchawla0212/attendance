import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import re
import zipfile

st.set_page_config(page_title="KollegeApply Attendance Summary Generator")
st.title("📋 KollegeApply Attendance Summary Generator")
st.caption("Upload Attendance Excel File (.xlsx or .xls)")

uploaded_file = st.file_uploader("Upload file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        buffer = BytesIO(uploaded_file.read())
        if not zipfile.is_zipfile(buffer):
            st.error("Uploaded file is not a valid Excel .xlsx file. Please re-save it in Excel.")
            st.stop()

        buffer.seek(0)

        # Load header from row 5 (index 4) for formatted version
        df_excel = pd.read_excel(buffer, sheet_name=0, header=4, engine="openpyxl")
        df_excel = df_excel.loc[:, ~df_excel.columns.str.contains('^Unnamed')]
        df_excel.columns = df_excel.columns.str.strip()

        required_cols = {"E. Code", "Name", "InTime", "OutTime", "Status"}

        if required_cols.issubset(set(df_excel.columns)):
            st.success("Detected new Daily Attendance format")

            buffer.seek(0)
            df_raw = pd.read_excel(buffer, sheet_name=0, header=None, engine="openpyxl")
            raw_date = df_raw.iloc[1, 4]  # Date at 2nd row, 5th column

            if isinstance(raw_date, datetime):
                full_date = raw_date.strftime("%Y-%m-%d")
            elif isinstance(raw_date, str):
                full_date = datetime.strptime(raw_date.strip(), "%d-%b-%Y").strftime("%Y-%m-%d")
            else:
                full_date = datetime.today().strftime("%Y-%m-%d")

            df_summary = df_excel[["E. Code", "Name", "InTime", "OutTime", "Status"]].copy()
            df_summary["Date"] = full_date
            df_summary.rename(columns={"E. Code": "Emp Code"}, inplace=True)
            df_summary = df_summary[["Emp Code", "Name", "Date", "InTime", "OutTime", "Status"]]

        else:
            st.success("Detected raw parser format")
            df_clean = pd.read_excel(buffer, engine="openpyxl")
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

                except Exception as e:
                    st.warning(f"Skipping row {i} due to error: {e}")
                    continue

            df_summary = pd.DataFrame(records)

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
