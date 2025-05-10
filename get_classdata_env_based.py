import re
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from dotenv import load_dotenv
import os

load_dotenv()  # 載入 .env 檔案

EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
LOGIN_URL = os.getenv("LOGIN_URL")  

def generate_progress_txt_from_class_mem(class_mem, output_path="progress_today.txt"):
    output_lines = ["本日進度："]
    for _, row in class_mem.iterrows():
        name = row['學生姓名']
        current = row['目前進度']
        last_stage = row['最後上課階段']
        m = re.search(r"L(\d+)", current)
        next_level = f"L{int(m.group(1)) + 1}" if m else current
        if "尚未完成" in last_stage:
            line = f"{name} 先完成 {current} 接著完成 {next_level}"
        else:
            line = f"{name} {next_level}"
        output_lines.append(line)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))
    print(f"✅ 已生成文字檔：{output_path}")
    return output_path

def generate_student_feedback_from_excel(file_path="student_progress.xlsx", output_txt="student_feedback.txt"):
    df = pd.read_excel(file_path)
    output_lines = ["【學生評語建議】"]
    for _, row in df.iterrows():
        name = str(row.get("學生姓名", "")).strip()
        progress = str(row.get("目前進度", "")).strip()
        status = str(row.get("本周是否完成", "")).strip()
        m = re.search(r"L(\d+)", progress)
        lesson = m.group(1) if m else "?"
        if status == '1':
            line = f"{name} 今天有順利完成第 {lesson} 課的專案！老師請 {name} 講解做法時，他/她能清楚說明，表現很不錯，繼續保持喔！"
        elif status == '2':
            line = f"{name} 今天的第 {lesson} 課專案尚未完成，目前完成到某階段，老師有確認目前進度沒有太大問題，下週再接再厲，加油！"
        elif status == '3':
            prev_lesson = str(int(lesson) - 1) if lesson.isdigit() else "?"
            line = f"{name} 今天先完成了上週的第 {prev_lesson} 課，本週的第 {lesson} 課已開始進行，狀況良好，下次再繼續努力！"
        else:
            continue
        output_lines.append(line)
    with open(output_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))
    print(f"✅ 已輸出學生評語建議到 {output_txt}")
    return output_txt

def main():
    print("請選擇使用情境：")
    print("1. 剛上課：登入、抓資料、生成進度表與建議")
    print("2. 下課評量：讀取現有 Excel 並生成評語")
    mode = input("請輸入模式代號（1 或 2）：").strip()
    if mode == '1':
        course_url = input("請貼上課程網址：").strip()
        options = Options()
        options.add_argument('--start-maximized')
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        try:
            driver.get(LOGIN_URL)
            time.sleep(2)
            driver.find_element(By.ID, "employee_email").send_keys(EMAIL)
            driver.find_element(By.ID, "employee_password").send_keys(PASSWORD)
            driver.find_element(By.NAME, "commit").click()
            time.sleep(3)
            if "sign_in" in driver.current_url:
                raise Exception("❌ 登入失敗，請檢查帳號密碼")
            print("✅ 登入成功")
            driver.get(course_url)
            time.sleep(2)
            soup = BeautifulSoup(driver.page_source, "lxml")
            frame = soup.find("turbo-frame", id="lesson_reports")
            if not frame or not frame.has_attr("src"):
                raise Exception("❌ 找不到 lesson_reports 的 URL")
            lesson_url = "https://corp.orangeapple.co" + frame["src"]
            print("✅ 取得課堂記錄頁面 URL：", lesson_url)
            driver.get(lesson_url)
            time.sleep(3)
            html = driver.page_source
            with open("lesson_reports_after_login.html", "w", encoding="utf-8") as f:
                f.write(html)
            print("📄 HTML 已儲存")
        finally:
            driver.quit()

        with open("lesson_reports_after_login.html", "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f, "lxml")

        progress_map = {}
        student_ths = soup.select('th[style=""]')
        for th in student_ths:
            full_text = th.get_text(" ", strip=True)
            match = re.search(r"([^\s]+)\s+L(\d+)", full_text)
            if match:
                name = match.group(1).strip()
                level = f"L{match.group(2)}"
                progress_map[name] = level

        results = []
        student_trs = soup.select("table#reports tbody tr")
        for tr in student_trs:
            name_span = tr.select_one("td span")
            if not name_span:
                continue
            student_name = name_span.text.strip()
            stage_div = tr.select_one("div.text-danger[data-bs-title='最後上課階段']")
            last_stage = stage_div.text.strip() if stage_div else "N/A"
            level = progress_map.get(student_name, "N/A")
            results.append({
                "學生姓名": student_name,
                "最後上課階段": last_stage,
                "目前進度": level,
                "本周是否完成": ""
            })

        df = pd.DataFrame(results)
        class_mem = df[
            df['目前進度'].notna() &
            (df['目前進度'].str.strip() != "") &
            (df['目前進度'].str.upper() != "N/A")
        ]
        df.to_excel("student_progress.xlsx", index=False)
        print("✅ 匯出完成：student_progress.xlsx")
        generate_progress_txt_from_class_mem(class_mem)
    elif mode == '2':
        generate_student_feedback_from_excel("student_progress.xlsx", "student_feedback.txt")
    else:
        print("❌ 無效的輸入，請輸入 1 或 2")

if __name__ == "__main__":
    main()
