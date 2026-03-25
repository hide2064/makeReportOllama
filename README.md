# makeReportOllama

Excel 売上データと PPTX テンプレートから、ローカル LLM (Ollama) を使って PowerPoint 報告書を自動生成する Web アプリです。

## 動作環境

| 項目 | 要件 |
|------|------|
| OS | Windows 10/11 |
| Python | 3.9 以上 |
| Node.js | 18 以上 |
| Ollama | バックグラウンドで起動済み (`http://localhost:11434`) |
| GPU | 不要（CPU 推論） |

## クイックスタート

```bat
start.bat
```

`start.bat` をダブルクリックするだけで以下が自動実行されます。

1. Node.js の存在確認（未インストール時は winget でインストール）
2. Python 仮想環境の構築・依存パッケージのインストール
3. フロントエンドの `npm install`
4. バックエンド (FastAPI) の起動確認・起動
5. フロントエンド (Vite) の起動確認・起動
6. ブラウザで `http://localhost:5173` を自動で開く

## 手動起動

### バックエンド

```bat
.venv\Scripts\activate
cd backend
uvicorn main:app --reload --port 8000
```

### フロントエンド

```bat
cd frontend
npm run dev
```

## テスト用ダミーデータの生成

```bat
.venv\Scripts\python setup_mock.py
```

`data/sales_data.xlsx` と `data/template.pptx` が生成されます。

## テスト実行

```bat
rem バックエンド (pytest)
.venv\Scripts\pytest tests/backend/ -v

rem フロントエンド (Vitest)
cd frontend
npm test
```

## ディレクトリ構成

```
makeReportOllama/
├── backend/               # FastAPI バックエンド
│   ├── main.py
│   ├── routers/report.py  # /api/generate エンドポイント
│   ├── services/
│   │   ├── excel_reader.py
│   │   ├── ollama_client.py
│   │   └── pptx_generator.py
│   └── requirements.txt
├── frontend/              # React + Vite フロントエンド
│   └── src/
├── data/                  # テスト用ダミーファイル
├── tests/                 # テストコード
├── .vscode/launch.json    # VSCode デバッグ設定
├── setup_mock.py          # ダミーデータ生成スクリプト
└── start.bat              # 全自動起動スクリプト
```
