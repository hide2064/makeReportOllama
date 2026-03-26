# 詳細設計書 — makeReportOllama

**バージョン**: 2.0.0
**最終更新**: 2026-03-27
**対象フェーズ**: Phase 1〜4 実装済み

---

## 目次

1. [システム概要](#1-システム概要)
2. [システムアーキテクチャ](#2-システムアーキテクチャ)
3. [技術スタック](#3-技術スタック)
4. [ディレクトリ構成](#4-ディレクトリ構成)
5. [バックエンド設計](#5-バックエンド設計)
6. [フロントエンド設計](#6-フロントエンド設計)
7. [AIパイプライン設計](#7-aiパイプライン設計)
8. [RAG設計](#8-rag設計)
9. [レポート生成設計](#9-レポート生成設計)
10. [API仕様](#10-api仕様)
11. [データモデル](#11-データモデル)
12. [処理フロー](#12-処理フロー)
13. [エラーハンドリング](#13-エラーハンドリング)
14. [テスト方針](#14-テスト方針)
15. [非機能要件](#15-非機能要件)
16. [起動・セットアップ](#16-起動セットアップ)
17. [フェーズ履歴](#17-フェーズ履歴)

---

## 1. システム概要

### 目的

Excel / CSV 形式の売上データと PowerPoint テンプレートをアップロードすると、ローカル LLM (Ollama) が売上分析を行い、日本語ビジネス文書として整形した PowerPoint 報告書を自動生成する Web アプリケーション。すべての処理はローカルで完結し、外部 API への送信は一切行わない。

### 主な機能

| 機能 | 説明 | フェーズ |
|------|------|---------|
| 売上データ読込 | Excel (.xlsx) / CSV (.csv) を自動判別・集計 | Phase 1 |
| AI 分析 (Analyst) | 数値データを構造化 JSON に変換 (qwen2.5:3b) | Phase 2 |
| AI 文書生成 (Writer) | 構造化データから日本語ビジネス文章を生成 (qwen3:8b) | Phase 2 |
| 進捗リアルタイム表示 | 生成ステップをポーリングで表示 | Phase 1 |
| RAG 参照 | 過去レポートを検索して文章品質・文脈一貫性を向上 | Phase 3 |
| 過去資料管理 UI | 過去報告書 PPTX のアップロード / 一覧 / 削除 | Phase 3 |
| 売上表スライド自動追加 | 商品×四半期クロス集計表をスライドに追加 | Phase 4a |
| グラフスライド自動追加 | 月次推移・商品構成グラフをスライドに追加 | Phase 4b |

### 動作環境

| 項目 | 要件 |
|------|------|
| OS | Windows 10 / 11 |
| Python | 3.10 以上 |
| Node.js | 18 以上 |
| Ollama | ローカルサーバー (http://localhost:11434) |
| 外部通信 | なし（完全ローカル動作） |

---

## 2. システムアーキテクチャ

```
┌────────────────────────────────────────────────────────────┐
│                    ユーザーブラウザ                          │
│             React SPA (http://localhost:5173)               │
│                                                            │
│  ┌─────────────────┐  ┌────────────────────────────────┐  │
│  │   UploadForm    │  │       ReferenceManager         │  │
│  │ (Excel + PPTX)  │  │  (過去レポート PPTX 登録/管理) │  │
│  └────────┬────────┘  └──────────────┬─────────────────┘  │
└───────────┼──────────────────────────┼────────────────────┘
            │ fetch (HTTP)             │ fetch (HTTP)
            ▼                         ▼
┌────────────────────────────────────────────────────────────┐
│           FastAPI Backend  (http://localhost:8000)          │
│                                                            │
│  ┌──────────────────────┐  ┌───────────────────────────┐  │
│  │  routers/report.py   │  │  routers/references.py    │  │
│  │  POST /api/generate  │  │  POST /api/references/    │  │
│  │  GET  /api/progress  │  │       upload              │  │
│  │  GET  /api/download  │  │  GET  /api/references     │  │
│  └──────────┬───────────┘  │  DELETE /api/references/  │  │
│             │              │       {file_id}            │  │
│             │              └─────────────┬─────────────┘  │
│             ▼                            ▼                 │
│  ┌──────────────────────────────────────────────────────┐  │
│  │                    Services Layer                     │  │
│  │  ┌──────────────┐ ┌──────────────┐ ┌─────────────┐  │  │
│  │  │ excel_reader │ │ollama_client │ │  rag_store  │  │  │
│  │  │  (pandas)    │ │  Analyst AI  │ │ (ChromaDB)  │  │  │
│  │  └──────────────┘ │  Writer AI   │ └──────┬──────┘  │  │
│  │  ┌──────────────┐ └──────┬───────┘        │         │  │
│  │  │pptx_generator│        │                │         │  │
│  │  │ (python-pptx │        │                │         │  │
│  │  │ +matplotlib) │        │                │         │  │
│  │  └──────────────┘        │                │         │  │
│  └──────────────────────────┼────────────────┼─────────┘  │
└─────────────────────────────┼────────────────┼────────────┘
                              │                │
               ┌──────────────┘                │
               ▼                               ▼
┌──────────────────────┐         ┌─────────────────────────┐
│   Ollama サーバー     │         │   ChromaDB (ローカル)    │
│  localhost:11434     │         │   data/chroma_db/        │
│  ・qwen2.5:3b        │         │   コサイン類似度検索      │
│  ・qwen3:8b          │         └─────────────────────────┘
│  ・nomic-embed-text  │
└──────────────────────┘
```

---

## 3. 技術スタック

### バックエンド

| カテゴリ | ライブラリ | バージョン | 用途 |
|----------|-----------|-----------|------|
| Web フレームワーク | FastAPI | ≥0.115 | REST API サーバー |
| ASGI サーバー | Uvicorn | ≥0.34 | アプリケーションサーバー |
| データ処理 | pandas | ≥2.2 | CSV / Excel 読込・集計 |
| Excel 読込 | openpyxl | ≥3.1 | .xlsx ファイル対応 |
| PPTX 生成 | python-pptx | ≥1.0 | PowerPoint 生成・テンプレート処理 |
| グラフ生成 | matplotlib | ≥3.8 | 売上グラフ画像生成（Agg バックエンド） |
| HTTP 通信 | httpx | ≥0.28 | Ollama API 通信（ストリーミング対応） |
| ファイルアップロード | python-multipart | ≥0.0.20 | multipart/form-data 処理 |
| ベクター DB | chromadb | ≥1.0 | 過去レポートのベクター保存・検索 |

### フロントエンド

| カテゴリ | ライブラリ | バージョン | 用途 |
|----------|-----------|-----------|------|
| UI フレームワーク | React | ^19.2.4 | SPA 構築 |
| ビルドツール | Vite | ^8.0.1 | バンドル・開発サーバー |
| 言語 | TypeScript | ~5.9.3 | 型安全な開発 |
| テスト | Vitest | ^4.1.1 | ユニット・コンポーネントテスト |
| テスト | @testing-library/react | ^16.3.2 | React コンポーネントテスト |

### AI・LLM

| モデル | エンジン | 役割 |
|--------|---------|------|
| qwen2.5:3b | Ollama | Analyst AI: 数値データ → 構造化 JSON 変換 |
| qwen3:8b | Ollama | Writer AI: 構造化 JSON → 日本語ビジネス文章生成 |
| nomic-embed-text | Ollama | 埋め込みモデル: 過去レポートのベクター化（RAG） |

---

## 4. ディレクトリ構成

```
makeReportOllama/
├── backend/
│   ├── main.py                      # FastAPI エントリポイント・CORS・ロギング
│   ├── requirements.txt             # Python 依存パッケージ
│   ├── app.log                      # アプリケーションログ（実行時生成）
│   ├── routers/
│   │   ├── report.py                # レポート生成エンドポイント群
│   │   └── references.py            # 過去資料管理エンドポイント群
│   └── services/
│       ├── excel_reader.py          # Excel / CSV 読込・集計サービス
│       ├── ollama_client.py         # Ollama API 通信（Analyst + Writer）
│       ├── pptx_generator.py        # PPTX 生成（テンプレート置換・表・グラフ）
│       └── rag_store.py             # ChromaDB ベクター管理サービス
│
├── frontend/
│   ├── package.json
│   ├── vite.config.ts
│   └── src/
│       ├── App.tsx                  # メインコンポーネント（状態管理・フロー制御）
│       ├── App.css                  # スタイル定義
│       ├── main.tsx                 # React エントリポイント
│       ├── components/
│       │   ├── UploadForm.tsx       # Excel / PPTX アップロードフォーム
│       │   ├── LoadingOverlay.tsx   # 生成中オーバーレイ（ステップ表示）
│       │   └── ReferenceManager.tsx # 過去資料管理 UI
│       └── test/
│           └── App.test.tsx         # フロントエンドテスト（Vitest）
│
├── tests/
│   └── backend/
│       ├── test_api.py              # FastAPI エンドポイント結合テスト
│       └── test_rag.py              # RAG・参照資料 ユニットテスト
│
├── data/
│   ├── template.pptx                # デフォルト PPTX テンプレート
│   ├── template_consultant.pptx     # コンサルタントスタイルテンプレート
│   ├── sales_data.xlsx / .csv       # サンプル売上データ（基本）
│   ├── sample_advanced.xlsx / .csv  # サンプル売上データ（利益率付き・2022〜2024）
│   ├── sample_report_2020Q1.pptx    # 過去レポートサンプル（RAG 用）
│   ├── sample_report_2021Q1.pptx
│   ├── sample_report_2022Q1.pptx
│   ├── sample_report_2023Q1.pptx
│   ├── sample_report_2024Q1.pptx
│   └── chroma_db/                   # ChromaDB 永続化ディレクトリ（実行時生成）
│
├── output/                          # 生成レポート出力先（実行時生成）
│   └── report.pptx
│
├── start.bat                        # Windows 一括起動スクリプト
├── CLAUDE.md                        # 開発メモ・次フェーズ計画
├── DESIGN.md                        # 本設計書
├── setup_mock.py                    # テスト用モックデータ生成
├── create_samples_advanced.py       # 高度サンプルデータ生成
├── create_sample_report.py          # 2024Q1 サンプルレポート生成
└── create_past_reports.py           # 2020〜2023 過去レポート一括生成
```

---

## 5. バックエンド設計

### 5.1 main.py

| 設定項目 | 値 |
|---------|-----|
| API タイトル | makeReportOllama API |
| バージョン | 1.0.0 |
| CORS 許可オリジン | http://localhost:5173, http://127.0.0.1:5173 |
| ログ出力先 | `backend/app.log`（UTF-8）＋標準出力 |
| ログフォーマット | `%(asctime)s [%(levelname)s] %(name)s: %(message)s` |

### 5.2 excel_reader.py

**入力**: Excel (.xlsx) または CSV (.csv) ファイルパス

**必須列**:

| 列名 | 型 | 説明 |
|------|----|------|
| 日付 | 日付文字列 | 売上日 (例: 2025-01-05) |
| 商品名 | 文字列 | 商品・サービス名 |
| 担当者 | 文字列 | 営業担当者名 |
| 地域 | 文字列 | 販売地域 |
| 数量 | 数値 | 販売数量 |
| 売上金額 | 数値 | 売上金額（円） |

**オプション列**（存在する場合に利益率グラフを自動生成）:

| 列名 | 説明 |
|------|------|
| 原価 | 原価金額 |
| 利益額 | 利益金額（月次利益率グラフに使用） |
| 利益率(%) | 利益率 |

**文字コード対応**: CSV は UTF-8 → Shift-JIS の順でフォールバック

**返却 dict**:

| キー | 型 | 説明 |
|-----|----|------|
| `total_amount` | int | 総売上金額 |
| `total_qty` | int | 総販売数量 |
| `period` | str | 集計期間文字列 |
| `raw_summary` | str | Ollama プロンプト用テキスト集計 |
| `by_product` | str | 商品別売上テキスト |
| `by_region` | str | 地域別売上テキスト |
| `by_rep` | str | 担当者別売上テキスト |
| `monthly_totals` | dict | 直近 12 ヶ月の月次合計 `{月: 金額}` |
| `product_totals` | dict | 商品別合計（降順） `{商品名: 金額}` |
| `quarterly_product_pivot` | DataFrame | 四半期×商品クロス集計（直近 8 四半期） |
| `monthly_margin` | dict\|None | 月次利益率 `{月: %}` ※利益額列がある場合のみ |

### 5.3 ollama_client.py

#### モデル・タイムアウト設定

| 定数 | 値 | 説明 |
|------|----|------|
| `MODEL_ANALYST` | `qwen2.5:3b` | Analyst AI |
| `MODEL_WRITER` | `qwen3:8b` | Writer AI |
| `OLLAMA_URL` | `http://localhost:11434/api/generate` | エンドポイント |
| `REQUEST_TIMEOUT` | 1200 秒（20 分） | read タイムアウト |

#### generate() 関数

- `stream: True` でトークンを逐次受信（サイレントハング防止）
- `think: False` で qwen3 の thinking モードを無効化
- 30 トークンごとに `on_token(count)` コールバックを呼び出し（進捗表示用）
- `<think>…</think>` ブロックをレスポンスから自動除去
- `httpx.Timeout(connect=30s, read=1200s, write=30s, pool=5s)`

#### Analyst AI プロンプト（英語指示 + 日本語 JSON 出力）

```
You are a data analyst. Output ONLY a JSON object.

{
  "period":          "集計期間",
  "total_sales":     総売上(数値),
  "total_qty":       総数量(数値),
  "top_products":    [{"name": ..., "amount": ...}],
  "bottom_products": [...],
  "top_regions":     [...],
  "bottom_regions":  [...],
  "top_reps":        [...],
  "key_facts":       ["重要な数値の事実"],
  "concerns":        ["懸念点・低迷要因"]
}

Sales data:
{raw_summary}
```

#### Writer AI プロンプト構造

```
【分析データ (JSON)】
{analyst_data}

【過去レポートからの参考情報】（RAG コンテキストがある場合のみ）
{rag_context}

【出力形式】
---SUMMARY---
（売上サマリー 300 字程度）
---ANALYSIS---
（課題・所見と改善策 300 字程度）
```

### 5.4 pptx_generator.py

#### 処理概要

1. テンプレート PPTX を読み込み
2. 全スライドのプレースホルダーをテキスト置換
3. 売上表スライドを末尾に追加（Phase 4a）
4. グラフスライドを末尾に追加（Phase 4b）
5. `output/report.pptx` に保存

#### プレースホルダー一覧

| プレースホルダー | 置換内容 |
|----------------|---------|
| `{{report_title}}` | 月次売上報告書（{period}） |
| `{{report_date}}` | 作成日: {YYYY年MM月DD日} |
| `{{summary_text}}` | Writer AI 生成サマリー |
| `{{analysis_text}}` | Writer AI 生成分析文 |

#### 売上表スライド（Phase 4a）

| 項目 | 仕様 |
|------|------|
| スライドサイズ | 16:9 (13.33" × 7.5") |
| 行構成 | ヘッダー行 / 商品行×N / 合計行 |
| 列構成 | 商品名 / 四半期ラベル×N（直近 8 四半期まで）/ 合計 |
| 数値フォーマット | X,XXX万（万円単位） |
| ヘッダー色 | Navy (#1B2E4C)、文字 White |
| 合計列色 | Gold (#C4973E)、文字 White |
| 合計行色 | Navy2 (#2C4A7A)、文字 White |
| 偶数データ行 | LightBlue (#E8EEF5) |

#### グラフスライド（Phase 4b）

| グラフ | 位置 | 種類 | データ |
|--------|------|------|--------|
| 月次売上推移 | 左 60% | 縦棒 | 直近 12 ヶ月の月次売上 |
| 商品別売上構成 | 右 40% | 横棒 | 商品別合計（上位 8 商品）|
| 利益率（オプション） | 左グラフ第 2 軸 | 折れ線 | 月次利益率（利益額列がある場合のみ）|

- matplotlib Agg バックエンドで PNG 生成 → python-pptx `add_picture()` で埋め込み
- 日本語フォント: Meiryo → MS Gothic → Yu Gothic → DejaVu Sans の優先順

### 5.5 rag_store.py

#### 定数

| 定数 | 値 |
|------|----|
| `CHROMA_DIR` | `data/chroma_db/` |
| `COLLECTION_NAME` | `past_reports` |
| `EMBED_MODEL` | `nomic-embed-text` |
| `OLLAMA_EMBED_URL` | `http://localhost:11434/api/embeddings` |
| `EMBED_TIMEOUT` | 60 秒 |
| `MAX_CHUNKS` | 5（検索で返す最大チャンク数） |
| `MAX_CTX_CHARS` | 1,500 文字（Writer に渡す上限） |
| `MIN_CHUNK_CHARS` | 30 文字（これ未満のチャンクは登録しない） |
| 類似度閾値 | 0.3（コサイン類似度；これ未満は除外） |

#### ChromaDB ドキュメント構造

| フィールド | 説明 |
|-----------|------|
| `id` | `{ファイル名MD5}_{チャンクインデックス}` |
| `document` | スライドテキスト（最大 300 字で検索結果に切り出し） |
| `embedding` | 768 次元 float 配列（nomic-embed-text 出力） |
| `metadata.filename` | 元 PPTX ファイル名 |
| `metadata.file_id` | ファイル名の MD5 ハッシュ（削除キー） |
| `metadata.chunk_idx` | スライドインデックス（0 始まり） |

---

## 6. フロントエンド設計

### 6.1 コンポーネント構成

```
App.tsx  ── 状態管理・API 呼び出し・フロー制御
├── LoadingOverlay.tsx   生成中オーバーレイ（ステップ文字表示）
├── UploadForm.tsx        Excel + PPTX ファイル選択フォーム
└── ReferenceManager.tsx  過去レポート管理（RAG 用）
```

### 6.2 App.tsx 状態定義

| 変数 | 型 | 説明 |
|------|----|------|
| `status` | `'idle'\|'loading'\|'success'\|'error'` | 全体状態 |
| `step` | `string` | ポーリングで取得した現在ステップメッセージ |
| `errorMsg` | `string` | エラーメッセージ |
| `downloadUrl` | `string` | 生成 PPTX の Blob オブジェクト URL |

#### 通信定数

| 定数 | 値 | 説明 |
|------|----|------|
| `POLL_INTERVAL` | 2,000 ms | ポーリング間隔 |
| `POLL_TIMEOUT` | 1,200,000 ms（20 分）| 最大待機時間 |

#### ポーリング耐障害設計

- ネットワーク瞬断（サーバー再起動等）は `catch` で無視して次のポーリングまで待機
- `AbortSignal.timeout(10_000)` で 1 回あたり 10 秒のタイムアウトを設定

### 6.3 ReferenceManager.tsx

| 操作 | エンドポイント | タイムアウト |
|------|--------------|------------|
| 一覧取得（マウント時） | GET /api/references | 10 秒 |
| PPTX 登録 | POST /api/references/upload | 120 秒 |
| 削除 | DELETE /api/references/{file_id} | 10 秒 |

---

## 7. AIパイプライン設計

### 2 モデルパイプライン（Phase 2）

```
売上データ (raw_summary テキスト)
         │
         ▼
┌─────────────────────┐
│   Analyst AI        │  qwen2.5:3b（小型・高速）
│   役割: 数値抽出     │  → 構造化 JSON のみ出力
└──────────┬──────────┘
           │ JSON
           ▼
┌─────────────────────┐
│   RAG 検索          │  ChromaDB（Phase 3）
│   過去文脈を取得     │  → 最大 1,500 字の参考テキスト
└──────────┬──────────┘
           │ JSON + 過去文脈
           ▼
┌─────────────────────┐
│   Writer AI         │  qwen3:8b（大型・高品質）
│   役割: 文章生成     │  → ---SUMMARY--- / ---ANALYSIS---
└──────────┬──────────┘
           │ テキスト解析
           ▼
     PPTX プレースホルダー埋め込み
```

### モデル選定根拠

| 比較軸 | Analyst (qwen2.5:3b) | Writer (qwen3:8b) |
|--------|---------------------|-------------------|
| タスク | 数値 → 構造化 JSON | JSON → 日本語ビジネス文章 |
| 要求品質 | 正確性・JSON 遵守 | 日本語流暢さ・文脈一貫性 |
| パラメータ数 | 3B（軽量） | 8B（高品質） |
| 推論速度 | 高速 | 低速（許容） |
| Think モード | 無効 (think: False) | 無効 (think: False) |

---

## 8. RAG設計

### 登録フロー

```
過去 PPTX ファイル
       │
       ▼ python-pptx
スライド単位でテキスト抽出
  ・{{...}} プレースホルダー行を除去
  ・30 文字未満のチャンクをスキップ
       │
       ▼ nomic-embed-text (Ollama)
768 次元浮動小数ベクター生成
  ・同一ファイルの再登録時は既存データを削除してから追加
       │
       ▼
ChromaDB に永続保存 (data/chroma_db/)
```

### 検索フロー（レポート生成時）

```
raw_summary テキスト
       │
       ▼ nomic-embed-text
クエリベクター（768 次元）
       │
       ▼ ChromaDB コサイン類似度検索（上位 5 件）
類似度 < 0.3 のチャンクを除外
       │
       ▼ 合計 1,500 字に達したら打ち切り
文脈テキスト文字列 → Writer AI プロンプトに挿入
（登録データ 0 件の場合はスキップして通常生成）
```

---

## 9. レポート生成設計

### 処理ステップと進捗メッセージ

| ステップ表示 | 内部処理 |
|------------|---------|
| [1/3] Excel / CSV を読み込んでいます... | `excel_reader.read_and_summarize()` |
| [2/3] 売上データを解析中... (Analyst: qwen2.5:3b) | Analyst AI 呼び出し |
| [2/3] 過去レポートから関連情報を検索中... (RAG) | `rag_store.search_context()` |
| [2/3] レポート文章を生成中... (Writer: qwen3:8b) | Writer AI 呼び出し |
| [3/3] PowerPoint レポートを生成しています... | `pptx_generator.generate_pptx()` |
| 完了しました！ | `done: True` をセット |

### 並列処理制御

| 設定 | 値 | 理由 |
|------|----|------|
| `ThreadPoolExecutor(max_workers=1)` | 1 スレッド | Ollama は単一リクエスト処理 |
| `threading.Lock()` | ステータス保護 | スレッドセーフな状態更新 |
| 二重投入防止 | `409 Conflict` | done=False かつ error='' のとき |

### 出力ファイル

- 保存先: `output/report.pptx`（固定パス、上書き保存）
- 生成 PPTX スライド構成:
  1. テンプレート由来のスライド（N 枚・プレースホルダー置換済み）
  2. 商品別売上表（Phase 4a で追加）
  3. 売上推移グラフ（Phase 4b で追加）

---

## 10. API仕様

### 10.1 レポート生成

#### POST /api/generate

**リクエスト**: `multipart/form-data`

| フィールド | 型 | 必須 | 説明 |
|-----------|----|------|------|
| `excel_file` | File | ○ | 売上データ (.xlsx / .csv) |
| `template_file` | File | ○ | PPTX テンプレート (.pptx) |

**レスポンス 200**:
```json
{ "status": "started" }
```

**エラー 409**: 処理中に再度リクエストされた場合

---

#### GET /api/progress

**レスポンス 200**:
```json
{
  "step":  "[2/3]  レポート文章を生成中です... (Writer: qwen3:8b)\n       生成中... 120 トークン生成済み",
  "done":  false,
  "error": ""
}
```

---

#### GET /api/download

**レスポンス 200**:
`Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation`

**エラー 404**: `output/report.pptx` が存在しない場合

---

### 10.2 参照資料管理

#### POST /api/references/upload

**リクエスト**: `multipart/form-data`

| フィールド | 型 | 説明 |
|-----------|----|------|
| `file` | File | 過去レポート (.pptx のみ受け付け) |

**レスポンス 200**:
```json
{ "filename": "report_2024Q1.pptx", "chunks": 5, "status": "registered" }
```

**エラー 400**: PPTX 以外のファイルが送信された場合
**エラー 500**: Ollama 接続失敗など

---

#### GET /api/references

**レスポンス 200**:
```json
{
  "references": [
    { "filename": "report_2024Q1.pptx", "file_id": "a3f2c1d8...", "chunks": 5 }
  ]
}
```

---

#### DELETE /api/references/{file_id}

**レスポンス 200**:
```json
{ "file_id": "a3f2c1d8...", "deleted_chunks": 5 }
```

**エラー 404**: 指定 `file_id` が存在しない場合

---

#### GET /health

**レスポンス 200**: `{ "status": "ok" }`

---

## 11. データモデル

### 売上 CSV フォーマット例

```csv
日付,商品名,担当者,地域,数量,売上金額,原価,利益額,利益率(%)
2024-01-10,プレミアムプラン,田中,東京,2,200000,90000,110000,55.0
2024-01-15,スタンダードプラン,佐藤,大阪,3,120000,54000,66000,55.0
```

### Analyst AI 出力 JSON スキーマ

```json
{
  "period":          "2024/01/01 ～ 2024/03/31",
  "total_sales":     38470000,
  "total_qty":       1250,
  "top_products":    [{ "name": "プレミアムプラン", "amount": 15200000 }],
  "bottom_products": [{ "name": "エントリーパッケージ", "amount": 2100000 }],
  "top_regions":     [{ "name": "東京", "amount": 12500000 }],
  "bottom_regions":  [{ "name": "九州", "amount": 1800000 }],
  "top_reps":        [{ "name": "田中", "amount": 8200000 }],
  "key_facts":       ["売上合計は前月比 +12.3%", "東京が全体の 32% を占める"],
  "concerns":        ["大阪地区が前月比 ▲8%", "エントリープランの単価下落"]
}
```

### Writer AI 出力テキスト形式

```
---SUMMARY---
今月の総売上は 3,847 万円となり、前月比 12.3% 増を達成しました。特にプレミアムプランが…

---ANALYSIS---
東京地区のプレミアムプランが好調な一方、大阪地区では前月比 8% の減少が見られます。…
```

---

## 12. 処理フロー

### レポート生成フロー

```
ユーザー          フロントエンド              バックエンド
   │                  │                          │
   ├─ファイル選択 ────▶│                          │
   ├─生成ボタン押下 ───▶│                          │
   │                  ├─ POST /api/generate ─────▶│ ← ファイル受信
   │                  │◀─ {status:"started"} ─────┤   バックグラウンド開始
   │                  │                          │
   │         [2秒待機] │                          ├─ [1/3] Excel 集計
   │                  ├─ GET /api/progress ───────▶│
   │                  │◀─ {step:"[1/3]..."} ───────┤
   │                  │                          │
   │   ステップ表示 ───│                          ├─ [2/3] Analyst AI
   │                  ├─ GET /api/progress ───────▶│
   │                  │◀─ {step:"[2/3]..."} ───────┤
   │                  │                          │
   │                  ├─ GET /api/progress ───────▶│ ← RAG 検索
   │                  │◀─ {step:"[2/3]..."} ───────┤
   │                  │                          │
   │                  ├─ GET /api/progress ───────▶│ ← Writer AI
   │                  │◀─ {step:"[2/3]..."} ───────┤
   │                  │                          │
   │                  ├─ GET /api/progress ───────▶│ ← PPTX 生成
   │                  │◀─ {done:true} ─────────────┤
   │                  │                          │
   │                  ├─ GET /api/download ───────▶│
   │                  │◀─ PPTX バイナリ ────────────┤
   │◀─ ダウンロード ───│                          │
```

### 過去資料登録フロー

```
ユーザー          フロントエンド              バックエンド         Ollama
   │                  │                          │                │
   ├─ PPTX 選択 ──────▶│                          │                │
   │                  ├─ POST /api/references/   │                │
   │                  │   upload ────────────────▶│                │
   │                  │                          ├─ PPTX テキスト抽出
   │                  │                          ├─ embed_text() ─▶│
   │                  │                          │◀─ 768 次元ベクター│
   │                  │                          ├─ ChromaDB 保存  │
   │                  │◀─ {chunks:N,"registered"}─┤                │
   │◀─ 登録完了表示 ───│                          │                │
```

---

## 13. エラーハンドリング

### バックエンドエラー処理方針

| エラー種別 | 発生箇所 | 対処 |
|----------|---------|------|
| ファイル未発見 | excel_reader | `FileNotFoundError` → progress.error にセット |
| 必須列不足 | excel_reader | `ValueError` → progress.error にセット |
| Ollama タイムアウト | ollama_client | `httpx.TimeoutException` → `RuntimeError` に変換 |
| Ollama 接続失敗 | ollama_client | `httpx.ConnectError` → `RuntimeError` に変換 |
| HTTP エラー | ollama_client | `httpx.HTTPStatusError` → `RuntimeError` に変換 |
| その他通信エラー | ollama_client | `httpx.HTTPError` → `RuntimeError` に変換 |
| Analyst JSON 解析失敗 | ollama_client | 空 dict にフォールバックして処理続行 |
| PPTX 生成失敗 | pptx_generator | `Exception` → progress.error にセット |
| PPTX 以外ファイル登録 | references | `HTTP 400` を返す |
| 埋め込み接続失敗 | rag_store | `RuntimeError("Ollama に接続できません...")` |

### フロントエンドエラー処理方針

| エラー種別 | 対処 |
|----------|------|
| POST /api/generate 失敗 | エラーメッセージ表示 → `status='error'` |
| ポーリング中の瞬断 | `catch` で無視してリトライ（無限リトライ可） |
| `progress.error` に値あり | エラーメッセージ表示 |
| 20 分タイムアウト | タイムアウトエラーを表示 |
| ダウンロード失敗 | エラーメッセージ表示 |

---

## 14. テスト方針

### バックエンドテスト

| ファイル | 対象 | Mock 対象 | テスト数 |
|---------|------|---------|---------|
| `tests/backend/test_api.py` | FastAPI エンドポイント | `routers.report.generate`（Ollama 通信） | 7 |
| `tests/backend/test_rag.py` | RAG ストア・参照 API | `_get_collection`、`embed_text`、`extract_chunks_from_pptx` | 13 |

#### test_api.py カバレッジ

| テスト名 | 内容 |
|---------|------|
| `test_health_endpoint` | /health が 200 を返す |
| `test_generate_success_excel` | Excel で Analyst + Writer 2 回呼び出し → done |
| `test_generate_success_csv` | CSV でも同様に done |
| `test_generate_ollama_error` | Analyst エラー → progress.error に入る |
| `test_generate_writer_error` | Writer エラー → progress.error に入る |
| `test_generate_invalid_excel` | 不正ファイル → progress.error に入る |
| `test_download_not_found` | report.pptx なし → 404 |

#### test_rag.py カバレッジ

| テストクラス / 関数 | 内容 |
|-----------------|------|
| `TestRagStore` | register / search（データあり・なし・低類似度）/ list / delete |
| `TestReferencesAPI` | upload 成功・非 PPTX 拒否・一覧取得・削除成功・削除 404 |
| `test_writer_prompt_includes_rag_context` | RAG あり時のプロンプト検証 |
| `test_writer_prompt_no_rag_context` | RAG なし時のプロンプト検証 |

### フロントエンドテスト

| ファイル | フレームワーク | テスト内容 |
|---------|-------------|---------|
| `frontend/src/test/App.test.tsx` | Vitest + Testing Library | 初期表示・ボタン無効・ポーリング成功・エラー表示 |

#### フロントエンド Mock 方針

- `global.fetch` を `vi.fn()` でモック
- `/api/references` は常に空リストを返す URL ルーティング式モック
- それ以外の URL はキューから順番に返す
- `vi.useFakeTimers()` + `vi.runAllTimersAsync()` でポーリングをテスト

---

## 15. 非機能要件

| 項目 | 仕様 |
|------|------|
| 外部通信 | なし（Ollama / ChromaDB はすべてローカル） |
| Ollama タイムアウト | 読み取り 1,200 秒（CPU 推論対応） |
| フロントエンドタイムアウト | 最大 20 分（1,200,000 ms） |
| グラフ生成 | matplotlib が未インストールの場合はグラフスライドをスキップ（フォールバック） |
| 二重処理防止 | 処理中の `/api/generate` 再投入に 409 を返す |
| ログ | `backend/app.log`（UTF-8）にすべての処理・エラーを記録 |
| CORS | localhost:5173 / 127.0.0.1:5173 のみ許可 |
| PPTX 出力 | `output/report.pptx` に固定保存（上書き） |
| ChromaDB 永続化 | `data/chroma_db/` に永続保存（Git 管理外） |

---

## 16. 起動・セットアップ

### 自動起動（start.bat）

`start.bat` をダブルクリックするだけで以下が自動実行される:

1. Node.js の確認・自動インストール（winget）
2. Ollama サービスの確認・起動
3. AI モデルの Pull（未取得の場合のみ）
   - `ollama pull qwen2.5:3b`
   - `ollama pull qwen3:8b`
   - `ollama pull nomic-embed-text`
4. Python 仮想環境 (`.venv/`) の作成と `pip install -r requirements.txt`
5. フロントエンド依存インストール（`npm install`）
6. バックエンド起動（uvicorn、port 8000）
7. フロントエンド起動（vite dev、port 5173）
8. ブラウザ自動起動（http://localhost:5173）

### 手動起動

```bash
# バックエンド
cd backend
../.venv/Scripts/python -m uvicorn main:app --host 0.0.0.0 --port 8000

# フロントエンド（別ターミナル）
cd frontend
npm run dev
```

### テスト実行

```bash
# バックエンドテスト
.venv/Scripts/python -m pytest tests/backend/ -v

# フロントエンドテスト
cd frontend
npm test
```

### サンプルデータ生成スクリプト

| スクリプト | 出力 | 用途 |
|---------|-----|------|
| `create_samples_advanced.py` | `data/sample_advanced.xlsx/.csv` | 利益率付き 2022〜2024 売上データ |
| `create_sample_report.py` | `data/sample_report_2024Q1.pptx` | RAG 用サンプルレポート 2024Q1 |
| `create_past_reports.py` | `data/sample_report_2020〜2023Q1.pptx` | RAG 用過去 4 年分レポート |

### ポート一覧

| サービス | ポート |
|---------|--------|
| フロントエンド（Vite） | 5173 |
| バックエンド（FastAPI） | 8000 |
| Ollama API | 11434 |

---

## 17. フェーズ履歴

| フェーズ | 主な実装内容 |
|---------|------------|
| **Phase 1** | 基本構成：単一モデル、ポーリング方式進捗表示、Excel / CSV 自動判別 |
| **Phase 2** | 2 モデル分離：Analyst (qwen2.5:3b) + Writer (qwen3:8b)、ストリーミング通信 |
| **Phase 3** | RAG 実装：ChromaDB + nomic-embed-text、過去資料管理 UI、文脈一貫性向上 |
| **Phase 4a** | 売上表スライド自動追加：商品×四半期クロス集計（python-pptx） |
| **Phase 4b** | グラフスライド自動追加：月次推移・商品構成グラフ（matplotlib）、利益率折れ線対応 |
