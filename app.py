import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import re

st.set_page_config(page_title="KollegeApply Attendance Summary Generator")
st.title("📋 KollegeApply Attendance Summary Generator")
st.caption("Upload Attendance Excel File (.xlsx or .xls)")

uploaded_file = st.file_uploader("Upload file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        file_ext = uploaded_file.name.split(".")[-1].lower()
        buffer = BytesIO(uploaded_file.read())
        buffer.seek(0)

        df_summary = None
        full_date = datetime.today().strftime("%Y-%m-%d")

        # Try case 1: "Original record" format
        try:
            df_log = pd.read_excel(buffer, engine="xlrd" if file_ext == "xls" else "openpyxl")
            if "Original record" in df_log.columns:
                st.success("Detected raw log format")

                df_clean = df_log.dropna(subset=['Original record']).reset_index(drop=True)
                date_line = str(df_clean.loc[0, 'Original record'])
                date_match = re.search(r'Date:(\d{4}-\d{1,2}-\d{1,2})', date_line)
                full_date = datetime.strptime(date_match.group(1), "%Y-%m-%d").strftime("%Y-%m-%d") if date_match else full_date

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

        except Exception:
            pass

        # Try case 2: New tabular format (header starts at row 5)
        if df_summary is None:
            buffer.seek(0)
            try:
                df_tabular = pd.read_excel(buffer, header=4, engine="xlrd" if file_ext == "xls" else "openpyxl")
                required_cols = {"E. Code", "Name", "InTime", "OutTime", "Status"}
                if required_cols.issubset(set(df_tabular.columns)):
                    st.success("Detected Daily Attendance tabular format")
                    buffer.seek(0)
                    df_head = pd.read_excel(buffer, header=None, engine="xlrd" if file_ext == "xls" else "openpyxl")
                    raw_date = df_head.iloc[1, 1]
                    if isinstance(raw_date, datetime):
                        full_date = raw_date.strftime("%Y-%m-%d")
                    elif isinstance(raw_date, str):
                        full_date = datetime.strptime(raw_date.strip(), "%d-%b-%Y").strftime("%Y-%m-%d")

                    df_summary = df_tabular[["E. Code", "Name", "InTime", "OutTime", "Status"]].copy()
                    df_summary.rename(columns={"E. Code": "Emp Code"}, inplace=True)
                    df_summary["Date"] = full_date
                    df_summary = df_summary[["Emp Code", "Name", "Date", "InTime", "OutTime", "Status"]]
            except Exception:
                pass

        if df_summary is not None and not df_summary.empty:
            st.dataframe(df_summary)
            csv = df_summary.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Download CSV", csv, file_name=f"summary_{full_date}.csv", mime='text/csv')
        else:
            st.error("❌ Could not detect format. Please upload a supported attendance Excel file.")

    except Exception as e:
        st.exception(e)
