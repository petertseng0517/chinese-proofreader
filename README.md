# 繁體中文 PDF 校稿工具

花蓮慈濟醫院公共傳播室內部工具：上傳 PDF，自動找出繁體中文錯別字，產出校稿報告並可下載 Excel。

## DEMO網址
https://chinese-proofreader.streamlit.app/

## 功能

1. 上傳 PDF 檔案（限制 20MB、100 頁以內）
2. 逐頁、逐行擷取文字（忽略圖片內容）
3. 呼叫 Claude API 進行繁體中文校稿
4. 網頁顯示校稿報告，格式：

   ```
   第X頁，第X行，「錯字」→ 建議「正確字」
   ```

   頁數採 **PDF 實際頁面順序**，非頁面上印刷的頁碼
5. 顯示統計：總字數、錯字數、錯誤率
6. 可下載 Excel 格式校稿報告（含「校稿結果」與「統計」兩個工作表），方便篩選過濾
7. 登入後才能使用（帳密驗證）

## 技術棧

- [Streamlit](https://streamlit.io/)：UI 與應用邏輯（單一 `app.py`）
- [PyMuPDF](https://pymupdf.readthedocs.io/)（`fitz`）：讀取 PDF 文字
- [Anthropic SDK](https://github.com/anthropics/anthropic-sdk-python)：呼叫 Claude API 校稿（`claude-opus-4-8`）
- [openpyxl](https://openpyxl.readthedocs.io/)：產生 Excel 校稿報告
- [python-dotenv](https://github.com/theskumar/python-dotenv)：讀取 `.env` 環境變數

## 安裝

```bash
# 建立並啟用虛擬環境
python3 -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate

# 安裝套件
pip install -r requirements.txt
```

## 環境變數設定

複製範例並編輯 `.env`（此檔案已加入 `.gitignore`，不會被提交到 GitHub）：

```bash
ANTHROPIC_API_KEY=sk-ant-...
APP_USERNAME=你的登入帳號
APP_PASSWORD_HASH=密碼的SHA-256雜湊值
```

`APP_PASSWORD_HASH` **不要存明文密碼**，請用以下指令產生雜湊值後填入：

```bash
python -c "import hashlib;print(hashlib.sha256('你的密碼'.encode()).hexdigest())"
```

## 執行

```bash
streamlit run app.py
```

開啟瀏覽器後先以 `.env` 中設定的帳密登入，再上傳 PDF 進行校稿。側邊欄有「登出」按鈕。

## 限制

- 檔案大小上限 20MB
- 頁數上限 100 頁
- 僅擷取文字，不處理圖片內容
- 目前僅支援單一帳號登入

## 部署注意事項

- 部署到雲端平台時，務必在平台的環境變數設定中加入 `ANTHROPIC_API_KEY`、`APP_USERNAME`、`APP_PASSWORD_HASH`，**不要**把 `.env` 檔案上傳到版本控制或部署套件中
- 登入狀態僅存於 Streamlit 的 session（記憶體），重新整理分頁或服務重啟後需重新登入
- 校稿結果由 Claude 模型產生，仍建議由人員複核後再正式採用
