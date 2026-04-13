# Vendas por Código — SC + SP

分析 SC + SP 銷售報表，自動加總 INATIVO 版本，輸出 CSV / XLSX。

## 本地測試

```bash
pip install -r requirements.txt
python app.py
# 打開 http://localhost:5000
```

## 部署到 Render（免費）

1. 把這個資料夾推到 GitHub（新建一個 repo）
2. 登入 https://render.com → New → Web Service
3. 連接你的 GitHub repo
4. 設定：
   - **Runtime**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
5. 點 Deploy — 幾分鐘後就有一個公開網址

## 檔案說明

- `app.py` — Flask 主程式（解析 Excel + API + 前端 HTML）
- `requirements.txt` — 依賴套件
- `render.yaml` — Render 自動部署設定
