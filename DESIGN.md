# 詳細設計書 — makeReportOllama

## 1. システム概要

| 項目 | 内容 |
|------|------|
| システム名 | makeReportOllama |
| 目的 | Excel 売上データと PPTX テンプレートから、ローカル LLM で PowerPoint 報告書を自動生成する |
| LLM | Ollama (http://localhost:11434) ※外部 API 不使用 |
| 実行環境 | Windows ネイティブ（Docker 不使用） |

---

## 2. アーキテクチャ

```
ブラウザ (React/Vite :5173)
    │
    │  POST /api/generate  (multipart/form-data)
    ▼
FastAPI バックエンド (:8000)
    │
    ├─ excel_reader.py   → pandas で Excel 集計
    ├─ ollama_client.py  → Ollama HTTP API 呼び出し
    └─ pptx_generator.py → python-pptx でテンプレート置換
```

### コンポーネント一覧

| コンポーネント | ファイル | 役割 |
|---------------|----------|------|
| API エントリポイント | `backend/main.py` | FastAPI アプリ初期化・CORS・ロギング設定 |
| レポート生成 API | `backend/routers/report.py` | `POST /api/generate` エンドポイント |
| Excel 読み込みサービス | `backend/services/excel_reader.py` | Excel 読み込み・集計・要約テキスト生成 |
| Ollama クライアント | `backend/services/ollama_client.py` | Ollama API 通信・プロンプト構築 |
| PPTX 生成サービス | `backend/services/pptx_generator.py` | テンプレートへのテキスト埋め込み・PPTX 保存 |
| UI メイン | `frontend/src/App.tsx` | ファイルアップロード・API 呼び出し・DL処理 |
| ファイルアップロードフォーム | `frontend/src/components/UploadForm.tsx` | Excel/PPTX 選択フォーム |
| ローディングオーバーレイ | `frontend/src/components/LoadingOverlay.tsx` | 処理中表示 |

---

## 3. API 仕様

### POST /api/generate

#### リクエスト

| パラメータ | 型 | 必須 | 説明 |
|-----------|-----|------|------|
| excel_file | File (multipart) | ○ | 売上データ `.xlsx` |
| template_file | File (multipart) | ○ | PPTX テンプレート `.pptx` |

#### レスポンス（正常時）

| 項目 | 値 |
|------|-----|
| Status | 200 |
| Content-Type | `application/vnd.openxmlformats-officedocument.presentationml.presentation` |
| Body | 生成された `report.pptx` バイナリ |

#### レスポンス（エラー時）

| Status | 原因 |
|--------|------|
| 400 | Excel の形式が不正、または必須列が不足 |
| 503 | Ollama への接続失敗またはタイムアウト |
| 500 | PPTX 生成中の予期しないエラー |

```json
// エラー時のボディ例
{ "detail": "エラー内容の説明テキスト" }
```

### GET /health

ヘルスチェック用エンドポイント。

```json
{ "status": "ok" }
```

---

## 4. データフロー

```
1. ユーザーが Excel + PPTX テンプレートをアップロード
2. FastAPI が一時ディレクトリに保存
3. excel_reader.py が pandas で集計
   - 商品別・地域別・担当者別売上を集計
   - Ollama プロンプト用の raw_summary テキストを生成
4. ollama_client.py が Ollama に 2 回リクエスト
   - summary_text: 売上サマリー（300字程度）
   - analysis_text: 所見・次月方針（300字程度）
5. pptx_generator.py がテンプレートのプレースホルダーを置換
   - {{report_title}}  → 月次売上報告書（集計期間）
   - {{report_date}}   → 作成日
   - {{summary_text}}  → Ollama 生成テキスト
   - {{analysis_text}} → Ollama 生成テキスト
6. 生成した PPTX をレスポンスとして返却
7. フロントエンドが Blob URL 経由でダウンロード
```

---

## 5. Excel ファイル仕様

### 必須列

| 列名 | 型 | 説明 |
|------|----|------|
| 日付 | 文字列 or 日付 | 売上日 (例: 2025-01-05) |
| 商品名 | 文字列 | 商品カテゴリ名 |
| 担当者 | 文字列 | 営業担当者名 |
| 地域 | 文字列 | 販売地域名 |
| 数量 | 数値 | 販売数量 |
| 売上金額 | 数値 | 売上金額（円） |

---

## 6. PPTX テンプレート仕様

テンプレート内のテキストボックスに以下のプレースホルダーを配置することで、レポートに値が埋め込まれます。

| プレースホルダー | スライド | 置換内容 |
|-----------------|----------|---------|
| `{{report_title}}` | 1 (表紙) | 報告書タイトル（集計期間入り） |
| `{{report_date}}` | 1 (表紙) | 作成日 |
| `{{summary_text}}` | 2 (サマリー) | Ollama 生成テキスト（売上サマリー） |
| `{{analysis_text}}` | 3 (所見) | Ollama 生成テキスト（所見・方針） |

---

## 7. 非機能要件

| 項目 | 内容 |
|------|------|
| タイムアウト | Ollama API タイムアウト: **6 分**（CPU 推論のため） |
| フロントエンドタイムアウト | `AbortSignal.timeout(360_000)` で 6 分 |
| ローディング表示 | Ollama 処理中はオーバーレイを表示し、ユーザーに処理中であることを伝える |
| ログ | `backend/app.log` にすべての処理・エラーを記録（UTF-8） |
| CORS | `http://localhost:5173` からのアクセスのみ許可 |

---

## 8. テスト方針

### バックエンド (pytest)

| テストファイル | 対象 | Ollama Mock |
|--------------|------|-------------|
| `tests/backend/test_excel_reader.py` | excel_reader | — |
| `tests/backend/test_pptx_generator.py` | pptx_generator | — |
| `tests/backend/test_api.py` | /api/generate エンドポイント | `unittest.mock.patch` でダミー返却 |

### フロントエンド (Vitest + Testing Library)

| テストファイル | 対象 |
|--------------|------|
| `frontend/src/test/App.test.tsx` | App コンポーネント（fetch をモック化） |

---

## 9. ディレクトリ構成

```
makeReportOllama/
├── backend/
│   ├── main.py
│   ├── routers/
│   │   └── report.py
│   ├── services/
│   │   ├── excel_reader.py
│   │   ├── ollama_client.py
│   │   └── pptx_generator.py
│   ├── requirements.txt
│   └── app.log              # 実行時生成
├── frontend/
│   ├── src/
│   │   ├── App.tsx
│   │   ├── App.css
│   │   ├── main.tsx
│   │   ├── index.css
│   │   ├── components/
│   │   │   ├── UploadForm.tsx
│   │   │   ├── UploadForm.css
│   │   │   ├── LoadingOverlay.tsx
│   │   │   └── LoadingOverlay.css
│   │   └── test/
│   │       ├── setup.ts
│   │       └── App.test.tsx
│   ├── package.json
│   └── vite.config.ts
├── data/
│   ├── sales_data.xlsx      # setup_mock.py で生成
│   └── template.pptx        # setup_mock.py で生成
├── output/                  # 生成された PPTX の保存先
├── tests/
│   └── backend/
│       ├── conftest.py
│       ├── test_excel_reader.py
│       ├── test_pptx_generator.py
│       └── test_api.py
├── .vscode/
│   └── launch.json
├── .venv/                   # Python 仮想環境
├── setup_mock.py
├── start.bat
├── README.md
└── DESIGN.md
```
