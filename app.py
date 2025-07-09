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
        df = pd.read_excel(uploaded_file, sheet_name=0)

        # Check if this is OLD FORMAT (Original record column exists)
        if "Original record" in df.columns:
            df_clean = df.dropna(subset=['Original record']).reset_index(drop=True)
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

        # Else: assume NEW FORMAT (structured sheet like you just uploaded)
        else:
            df_table = pd.read_excel(uploaded_file, sheet_name=0, header=4)
            meta_df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            raw_date = str(meta_df.iloc[1, 4]).strip()

            try:
                full_date = datetime.strptime(raw_date, "%d-%b-%Y").strftime("%Y-%m-%d")
            except:
                full_date = raw_date

            df_filtered = df_table[df_table["Status"].astype(str).str.lower() == "present"]
            df_filtered = df_filtered[["Name", "In Time", "Out Time", "Status"]]
            df_filtered["Date"] = full_date
            df_summary = df_filtered[["Name", "Date", "In Time", "Out Time", "Status"]]

            st.success(f"✅ Parsed {len(df_summary)} records from new format ({full_date})")
            st.dataframe(df_summary)

        # Download button
        buffer = BytesIO()
        df_summary.to_csv(buffer, index=False)
        buffer.seek(0)
        st.download_button(
            label="📥 Download Summary CSV",
            data=buffer,
            file_name=f"attendance_summary_{full_date}.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
