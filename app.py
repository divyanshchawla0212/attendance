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

def is_structured_format(df):
    return {"E. Code", "Name", "InTime", "OutTime", "Status"}.issubset(df.columns) or \
           {"E. Code", "Name", "In Time", "OutTime", "Status"}.issubset(df.columns)

if uploaded_file:
    try:
        file_ext = uploaded_file.name.split('.')[-1]
        buffer = BytesIO(uploaded_file.read())

        # Format 3: Old .xls (parsed without zip check)
        if file_ext == "xls":
            df_excel = pd.read_excel(buffer, sheet_name=0, header=4, engine="xlrd")
        else:
            # Format 1 or 2: .xlsx (verify zip)
            if not zipfile.is_zipfile(buffer):
                st.error("Uploaded file is not a valid .xlsx file. Please re-save it properly.")
                st.stop()
            buffer.seek(0)
            df_excel = pd.read_excel(buffer, sheet_name=0, header=4, engine="openpyxl")

        # Format 1 or 3: Structured
        if is_structured_format(df_excel):
            st.success("✅ Detected structured Daily Attendance format")

            buffer.seek(0)
            df_raw = pd.read_excel(buffer, sheet_name=0, header=None, engine="openpyxl" if file_ext == "xlsx" else "xlrd")
            raw_date = df_raw.iloc[1, 1]
            if isinstance(raw_date, datetime):
                full_date = raw_date.strftime("%Y-%m-%d")
            elif isinstance(raw_date, str):
                try:
                    full_date = datetime.strptime(raw_date.strip(), "%d-%b-%Y").strftime("%Y-%m-%d")
                except:
                    full_date = datetime.today().strftime("%Y-%m-%d")
            else:
                full_date = datetime.today().strftime("%Y-%m-%d")

            df_summary = df_excel[["E. Code", "Name", "In Time" if "In Time" in df_excel.columns else "InTime",
                                   "OutTime", "Status"]].copy()
            df_summary["Date"] = full_date
            df_summary.rename(columns={"E. Code": "Emp Code"}, inplace=True)
            df_summary = df_summary[["Emp Code", "Name", "Date", "In Time" if "In Time" in df_summary.columns else "InTime",
                                     "OutTime", "Status"]]

        else:
            # Format 2: Raw parser format (Original record)
            buffer.seek(0)
            st.success("✅ Detected raw log format")
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

        # Show & Download
        st.dataframe(df_summary)

        csv_buffer = BytesIO()
        df_summary.to_csv(csv_buffer, index=False)
        csv_buffer.seek(0)

        st.download_button(
            label=f"📥 Download CSV (summary_{full_date}.csv)",
            data=csv_buffer,
            file_name=f"attendance_summary_{full_date}.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
