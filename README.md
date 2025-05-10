# OA 課堂紀錄自動處理工具 🧑‍🏫

這是一個使用 Python + Selenium 的自動爬蟲腳本，可協助你完成以下任務：

- 登入 OA 系統
- 自動取得課堂紀錄 HTML
- 擷取學生進度與階段
- 匯出為 Excel 表格
- 生成本日學習進度文字檔
- 下課後根據填寫狀態，生成學生評語文字建議

---

## 🔧 使用方式

### 1️⃣ 安裝相依套件

```bash
pip install -r requirements.txt
```

### 2️⃣ 設定帳密資訊 `config.txt`

請建立一個名為 `config.txt` 的檔案（放在與程式同資料夾），內容格式如下：

```txt
EMAIL=your_email@example.com
PASSWORD=your_password
LOGIN_URL=https://corp.companylink.co/employees/sign_in
```

👉 若你想快速開始，可以複製 `config_sample.txt` 為範本。

---

## 🚀 執行主程式

```bash
python get_classdata_txt_config.py
```

程式會要求你選擇模式：

1. **剛上課模式**：會進行登入、抓取資料、產生進度 Excel、文字建議
2. **下課評量模式**：僅針對已有的 `student_progress.xlsx` 產出文字建議

---

## 📁 輸出結果

- `student_progress.xlsx`：學生出席與進度記錄
- `progress_today.txt`：根據進度推薦的今日目標
- `student_feedback.txt`：課後可用的評語建議文字

---

## ✅ 注意事項

- 程式會模擬瀏覽器登入，請勿濫用
- **請勿上傳 `config.txt` 到 GitHub！**
