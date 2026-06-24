# 繁體中文 PDF 校稿工具

內部工具：上傳 PDF → 逐頁逐行擷取文字（忽略圖片） → 呼叫 Claude API 找錯別字 → 網頁顯示校稿報告 → 下載 Excel 報告。

頁數一律採 PDF 實際頁面順序（第幾個 page object），不使用頁面上印刷的頁碼。

## 技術棧
- Streamlit（UI + 應用邏輯，單一 `app.py`）
- PyMuPDF (`fitz`)：讀取 PDF 文字（逐行），不處理圖片
- `anthropic` 官方 SDK：呼叫 Claude API 校對繁體中文錯別字
- `openpyxl`：產生可下載的 Excel 校稿報告

## 限制
- 檔案大小上限 20MB
- 頁數上限 100 頁

## API 設定
- `ANTHROPIC_API_KEY` 存放於專案根目錄 `.env`（已加入 `.gitignore`，不會上傳到 GitHub）
- `app.py` 用 `python-dotenv` 的 `load_dotenv()` 讀取 `.env`，再由 `anthropic.Anthropic()` 預設行為取用環境變數（不寫死金鑰）
- 模型：`claude-opus-4-8`，搭配 adaptive thinking + `output_config.effort: "high"`
- 校稿結果用 `output_config.format`（json_schema）強制結構化輸出，避免手動解析自由文字

## 登入驗證
部署在網路上時需要登入才能使用，帳密設定於 `.env`：
- `APP_USERNAME`：帳號（明文）
- `APP_PASSWORD_HASH`：密碼的 SHA-256 雜湊值（不存明文密碼）

產生密碼雜湊值：
```bash
./venv/bin/python -c "import hashlib;print(hashlib.sha256('你的密碼'.encode()).hexdigest())"
```
將輸出結果填入 `.env` 的 `APP_PASSWORD_HASH`。目前僅支援單一帳號；登入狀態存於 Streamlit `session_state`，重新整理頁面或重啟服務需重新登入。側邊欄有「登出」按鈕。

## 執行
```bash
pip install -r requirements.txt
# 編輯 .env，填入 ANTHROPIC_API_KEY、APP_USERNAME、APP_PASSWORD_HASH
streamlit run app.py
```
