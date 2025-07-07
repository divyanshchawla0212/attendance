import streamlit as st
import pandas as pd
from datetime import datetime
import re
from io import BytesIO

st.title("Attendance Parser and Summary Generator")

uploaded_file = st.file_uploader("Upload Excel File (.xls or .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df_clean = df.dropna(subset=['Original record']).reset_index(drop=True)

        # Extract date from first record (e.g., "Date:2025-7-3~2025-7-3")
        date_line = str(df_clean.loc[0, 'Original record'])
        date_match = re.search(r'Date:(\d{4}-\d{1,2}-\d{1,2})', date_line)
        if date_match:
            full_date = datetime.strptime(date_match.group(1), "%Y-%m-%d").strftime("%Y-%m-%d")
        else:
            st.error("Could not extract date from the first row.")
            st.stop()

        # Parse records
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

        # Show results
        df_summary = pd.DataFrame(records)
        st.success(f"Parsed {len(df_summary)} records for date {full_date}")
        st.dataframe(df_summary)

        # Download button
        csv_buffer = BytesIO()
        df_summary.to_csv(csv_buffer, index=False)
        csv_buffer.seek(0)

        st.download_button(
            label=f"Download CSV (df_summary_{full_date}.csv)",
            data=csv_buffer,
            file_name=f'df_summary_{full_date}.csv',
            mime='text/csv'
        )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
