import time
import pandas as pd
import streamlit as st
import tempfile
import os
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

st.set_page_config(page_title="KollegeApply Tools", layout="centered")
st.title("🧠 KollegeApply Tools: Index & Attendance Checker")

tool = st.sidebar.radio("Choose Tool", ["Google Index Checker", "Attendance Extractor"])

# ======================= 🔍 GOOGLE INDEX CHECKER =======================
if tool == "Google Index Checker":
    uploaded_file = st.file_uploader("📤 Upload a CSV with column 'URL'", type=["csv"])
    if uploaded_file:
        df = pd.read_csv(uploaded_file)
        if "URL" not in df.columns:
            st.error("❌ Your CSV must contain a column named 'URL'.")
            st.stop()

        urls = df["URL"].dropna().unique().tolist()
        st.info(f"✅ {len(urls)} URLs found.")

        if st.button("🚀 Start Google Index Check"):
            chrome_options = Options()
            # chrome_options.add_argument("--headless")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")

            driver = webdriver.Chrome(options=chrome_options)

            results = []
            progress = st.progress(0)
            status = st.empty()

            for i, url in enumerate(urls):
                status.text(f"🔍 Checking: {url}")
                try:
                    driver.get("https://www.google.com")
                    time.sleep(2)

                    try:
                        agree_btn = driver.find_element(By.XPATH, '//button[contains(text(), "I agree")]')
                        agree_btn.click()
                        time.sleep(1)
                    except:
                        pass

                    search_box = driver.find_element(By.NAME, "q")
                    search_box.clear()
                    search_box.send_keys(f"site:{url}")
                    search_box.send_keys(Keys.RETURN)
                    time.sleep(3)

                    # CAPTCHA check
                    if "sorry/index" in driver.current_url or "interstitial" in driver.current_url:
                        st.warning(f"⚠️ CAPTCHA triggered for: {url}")
                        st.info("🔓 Solve CAPTCHA in browser, then press Enter in terminal.")
                        input("▶️ Press Enter after solving CAPTCHA...")

                    found = False
                    links = driver.find_elements(By.XPATH, '//div[@id="search"]//a[@href]')
                    for link in links:
                        href = link.get_attribute("href")
                        if href and "kollegeapply.com" in href:
                            found = True
                            break

                    results.append({"URL": url, "Indexed on Google?": "Yes" if found else "No"})

                except Exception as e:
                    results.append({"URL": url, "Indexed on Google?": "Error"})

                progress.progress((i + 1) / len(urls))

            driver.quit()
            df_result = pd.DataFrame(results)
            st.success("✅ Done!")
            st.dataframe(df_result)

            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
            df_result.to_csv(tmp.name, index=False)
            with open(tmp.name, "rb") as f:
                st.download_button("📥 Download CSV", data=f, file_name="index_check.csv")
            os.unlink(tmp.name)

# ======================= 📅 ATTENDANCE EXTRACTOR =======================
elif tool == "Attendance Extractor":
    excel_file = st.file_uploader("📤 Upload Excel file (Daily Attendance Report)", type=["xlsx"])

    if excel_file:
        df = pd.read_excel(excel_file)
        df_clean = df.dropna(subset=['Original record']).reset_index(drop=True)

        # Extract date from row 0
        date_line = str(df_clean.loc[0, 'Original record'])
        date_match = re.search(r'Date:(\d{4}-\d{1,2}-\d{1,2})', date_line)
        if date_match:
            full_date = datetime.strptime(date_match.group(1), "%Y-%m-%d").strftime("%Y-%m-%d")
        else:
            st.error("❌ Could not extract date from the first row.")
            st.stop()

        records = []
        for i in range(1, len(df_clean) - 2, 3):
            try:
                info = df_clean.loc[i, 'Original record']
                times_raw = df_clean.loc[i + 2, 'Original record']

                name = re.search(r'Name:(.*?)Dept', info)
                name = name.group(1).strip() if name else "Unknown"

                dept = re.search(r'Dept\.:([^\s]+)', info)
                dept = dept.group(1).strip() if dept else "Unknown"

                times = [t.strip() for t in times_raw.strip().split('\n') if t.strip()]
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

        df_attendance = pd.DataFrame(records)
        st.success("✅ Attendance parsed successfully!")
        st.dataframe(df_attendance)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
        df_attendance.to_csv(tmp.name, index=False)
        with open(tmp.name, "rb") as f:
            st.download_button("📥 Download Attendance CSV", data=f, file_name="attendance_summary.csv")
        os.unlink(tmp.name)
