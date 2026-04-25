# 盤點行程異動聯絡單產生器 — 後端 API

## 部署到 Render（免費）

1. 前往 https://render.com 註冊免費帳號
2. 點擊「New」→「Web Service」
3. 連結這個 GitHub Repository
4. 設定：
   - Name: inventory-api
   - Runtime: Python 3
   - Build Command: pip install -r requirements.txt
   - Start Command: gunicorn app:app
5. 點擊「Create Web Service」
6. 等待部署完成（約 2 分鐘）
7. 取得 API 網址（格式：https://inventory-api-xxxx.onrender.com）
8. 把這個網址填入前端 index.html 的 API_URL 變數

## API 端點

- GET  /health         — 健康檢查
- POST /compare        — 比對兩個班表（回傳 JSON）
- POST /generate       — 比對並產出 Excel（回傳檔案）

## 注意事項
- Render 免費方案閒置 15 分鐘後會休眠，第一次呼叫需等約 30 秒喚醒
- template.xlsx 必須放在同一個資料夾
