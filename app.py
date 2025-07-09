import streamlit as st
import pandas as pd
from datetime import datetime
import re
from io import BytesIO

st.set_page_config(page_title="KollegeApply Attendance Parser", layout="centered")
st.title("📋 KollegeApply Attendance Summary Generator")

uploaded_file = st.file_uploader("Upload Attendance Excel File (.xlsx or .xls)", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read raw bytes and wrap in BytesIO for consistent engine use
        excel_bytes = uploaded_file.read()
        buffer = BytesIO(excel_bytes)

        # Load a preview dataframe to check format
        preview_df = pd.read_excel(buffer, sheet_name=0, nrows=2, header=None, engine="openpyxl")

        # OLD FORMAT CHECK (Original record column exists)
        buffer.seek(0)
        df_check = pd.read_excel(buffer, sheet_name=0, engine="openpyxl")
        if "Original record" in df_check.columns:
            df_clean = df_check.dropna(subset=['Original record']).reset_index(drop=True)
            date_line = str(df_clean.loc[0, 'Original record'])
            date_match = re.search(r'Date:(\d{4}-\d{1,2}-\d{1,2})', date_line)
            if date_match:
                full_date = datetime.strptime(date_match.group(1), "%Y-%m-%d").strftime("%Y-%m-%d")
            else:
                st.error("Could not extract date from the first row.")
                st.stop()

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
                    continue

            df_summary = pd.DataFrame(records)
            st.success(f"✅ Parsed {len(df_summary)} records from old format ({full_date})")
            st.dataframe(df_summary)

        # NEW FORMAT (structured sheet like '311 Daily Attendance Report')
        else:
            buffer.seek(0)
            df_table = pd.read_excel(buffer, sheet_name=0, header=4, engine="openpyxl")
            buffer.seek(0)
            meta_df = pd.read_excel(buffer, sheet_name=0, header=None, engine="openpyxl")

            try:
                raw_date = str(meta_df.iloc[1, 4]).strip()
                full_date = datetime.strptime(raw_date, "%d-%b-%Y").strftime("%Y-%m-%d")
            except:
                full_date = "Unknown"

            if "Name" in df_table.columns and "Status" in df_table.columns:
                df_filtered = df_table[df_table["Status"].astype(str).str.lower() == "present"]
                df_filtered = df_filtered[["Name", "In Time", "OutTime", "Status"]]
                df_filtered["Date"] = full_date

                df_summary = df_filtered.rename(columns={"OutTime": "Out Time"})[
                    ["Name", "Date", "In Time", "Out Time", "Status"]
                ]

                st.success(f"✅ Parsed {len(df_summary)} present records from new format ({full_date})")
                st.dataframe(df_summary)
            else:
                st.error("Could not find expected columns in the uploaded sheet.")
                st.stop()

        # ===== CSV Download =====
        csv_buffer = BytesIO()
        df_summary.to_csv(csv_buffer, index=False)
        csv_buffer.seek(0)
        st.download_button(
            label="📥 Download Attendance Summary CSV",
            data=csv_buffer,
            file_name=f"attendance_summary_{full_date}.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
