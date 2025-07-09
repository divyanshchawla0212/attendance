import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="KollegeApply Attendance Summary Generator", page_icon="📝")

st.title("📝 KollegeApply Attendance Summary Generator")
st.subheader("Upload Attendance Excel File (.xlsx or .xls)")

uploaded_file = st.file_uploader("Upload file", type=["xlsx", "xls"])

def detect_format(df):
    """Detects which report format is present: '311' (new) or '407' (old)."""
    cols = [col.lower() for col in df.columns]
    if 'e. code' in cols or 'e code' in cols:
        return '311'
    elif 'employee code' in cols or 'emp code' in cols:
        return '407'
    else:
        return 'unknown'

def extract_311_format(df, sheet):
    # Attempt to find header row (should be row 6 or nearby)
    for i in range(0, 10):
        row = df.iloc[i].astype(str).str.lower().tolist()
        if "e. code" in row or "e code" in row:
            df.columns = df.iloc[i]
            df = df[i + 1:].reset_index(drop=True)
            break

    # Clean up
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.rename(columns=lambda x: str(x).strip())
    df = df[["E. Code", "Name", "Shift", "InTime", "OutTime", "Work Dur", "OT", "Tot. Dur", "Status", "Remarks"]]
    df = df[df["E. Code"].notnull()]
    df["Date"] = extract_date_from_sheet(sheet)
    return df

def extract_407_format(df, sheet):
    # Find header row
    for i in range(0, 10):
        row = df.iloc[i].astype(str).str.lower().tolist()
        if "employee code" in row or "emp code" in row:
            df.columns = df.iloc[i]
            df = df[i + 1:].reset_index(drop=True)
            break

    df = df.loc[:, ~df.columns.duplicated()]
    df = df.rename(columns=lambda x: str(x).strip())

    df = df[["Employee Code", "Name", "Shift", "InTime", "OutTime", "Work Duration", "OT", "Total Duration", "Status", "Remarks"]]
    df = df[df["Employee Code"].notnull()]
    df["Date"] = extract_date_from_sheet(sheet)
    return df

def extract_date_from_sheet(sheet):
    try:
        for i in range(0, 10):
            row = sheet.iloc[i].astype(str).tolist()
            for item in row:
                if "202" in item:
                    dt = item.strip()
                    try:
                        return datetime.strptime(dt, "%d-%b-%Y").strftime("%d-%m-%Y")
                    except:
                        try:
                            return datetime.strptime(dt, "%d-%b-%y").strftime("%d-%m-%Y")
                        except:
                            try:
                                return datetime.strptime(dt, "%d-%m-%Y").strftime("%d-%m-%Y")
                            except:
                                continue
    except:
        pass
    return ""

if uploaded_file:
    try:
        sheet = pd.read_excel(uploaded_file, sheet_name=0, header=None)

        format_type = detect_format(sheet)
        if format_type == "311":
            final_df = extract_311_format(sheet.copy(), sheet)
        elif format_type == "407":
            final_df = extract_407_format(sheet.copy(), sheet)
        else:
            st.warning("⚠️ Unable to detect supported format. Please check the file structure.")
            st.stop()

        st.success("✅ File processed successfully!")
        st.write(final_df)

        csv = final_df.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Processed CSV", data=csv, file_name="attendance_summary.csv", mime='text/csv')

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
