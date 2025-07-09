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

def extract_date_from_header(df_raw):
    try:
        # Look for any cell with a date-like pattern
        for row in df_raw.iloc[:5].values.flatten():
            if isinstance(row, datetime):
                return row.strftime("%Y-%m-%d")
            if isinstance(row, str):
                try:
                    return datetime.strptime(row.strip(), "%d-%b-%Y").strftime("%Y-%m-%d")
                except:
                    pass
    except:
        pass
    return datetime.today().strftime("%Y-%m-%d")

if uploaded_file:
    try:
        buffer = BytesIO(uploaded_file.read())
        buffer.seek(0)

        extension = uploaded_file.name.split(".")[-1].lower()

        # Try all formats
        df_summary = None
        full_date = None

        if extension == "xls":
            df = pd.read_excel(buffer, engine="xlrd", header=None)
        else:
            df = pd.read_excel(buffer, engine="openpyxl", header=None)

        # Try Format 1: Standard MoM Format
        buffer.seek(0)
        try:
            df_clean = pd.read_excel(buffer, engine="openpyxl")
            if "Original record" in df_clean.columns:
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
                    except:
                        continue
                df_summary = pd.DataFrame(records)

        except Exception:
            pass

        # Try Format 2 or 3: Header-based Attendance Table
        if df_summary is None:
            buffer.seek(0)
            if extension == "xls":
                df_excel = pd.read_excel(buffer, engine="xlrd", header=5)
                df_raw = pd.read_excel(buffer, engine="xlrd", header=None)
            else:
                df_excel = pd.read_excel(buffer, engine="openpyxl", header=5)
                df_raw = pd.read_excel(buffer, engine="openpyxl", header=None)

            full_date = extract_date_from_header(df_raw)

            columns = df_excel.columns.str.strip().str.lower()
            df_excel.columns = columns

            # Find matching columns using fuzzy mapping
            possible_mappings = {
                "emp_code": ["e. code", "emp code", "empcode", "e code"],
                "name": ["name"],
                "in_time": ["in time", "intime", "in"],
                "out_time": ["out time", "outtime", "out"],
                "status": ["status"]
            }

            def find_column(possible_names):
                for name in possible_names:
                    for col in columns:
                        if name in col:
                            return col
                return None

            emp_col = find_column(possible_mappings["emp_code"])
            name_col = find_column(possible_mappings["name"])
            in_col = find_column(possible_mappings["in_time"])
            out_col = find_column(possible_mappings["out_time"])
            status_col = find_column(possible_mappings["status"])

            if None in [emp_col, name_col, in_col, out_col, status_col]:
                raise ValueError("Could not detect required columns.")

            df_summary = df_excel[[emp_col, name_col, in_col, out_col, status_col]].copy()
            df_summary.columns = ["Emp Code", "Name", "In Time", "Out Time", "Status"]
            df_summary["Date"] = full_date
            df_summary = df_summary[["Emp Code", "Name", "Date", "In Time", "Out Time", "Status"]]

        # Show and download
        st.success(f"Parsed Successfully for {full_date}")
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
        st.error(f"❌ Error: {e}")
