# makeReportOllama DB 仕様書

**バージョン**: 1.0
**作成日**: 2026-03-29

---

## 1. ストレージ概要

本システムはリレーショナルデータベースを使用しない。
永続化ストレージとして以下の 3 種類を使用する。

| ストレージ種別 | 用途 | 場所 |
|-------------|-----|------|
| JSON ファイル | 生成履歴の管理 | `output/history.json` |
| ChromaDB（ベクター DB） | RAG 参照チャンクの保存・検索 | `data/chroma_db/` |
| ファイルシステム | 生成済み PPTX ファイル | `output/report_{job_id}.pptx` |

---

## 2. history.json — 生成履歴ストア

### 2.1 概要

| 項目 | 内容 |
|-----|------|
| ファイルパス | `output/history.json` |
| フォーマット | JSON 配列 |
| 最大件数 | 50 件（超過時は最古エントリを自動削除） |
| アクセス制御 | スレッドロック（`threading.Lock`）により排他制御 |
| 文字コード | UTF-8 |

### 2.2 スキーマ

```json
[
  {
    "job_id":            "ab12cd34ef56",
    "created_at":        "2026-03-29T12:34:56Z",
    "original_filename": "sales_2024Q1.xlsx",
    "output_path":       "/abs/path/output/report_ab12cd34ef56.pptx",
    "analyst_model":     "qwen2.5:3b",
    "writer_model":      "qwen3:8b"
  }
]
```

### 2.3 フィールド定義

| フィールド名 | 型 | 必須 | 説明 |
|------------|----|----|------|
| job_id | string | ○ | UUID の先頭 12 文字（例: `ab12cd34ef56`） |
| created_at | string | ○ | ISO 8601 UTC 形式のタイムスタンプ |
| original_filename | string | ○ | アップロードされた元のファイル名 |
| output_path | string | ○ | 生成された PPTX の絶対パス |
| analyst_model | string | ○ | 使用した Analyst モデル名 |
| writer_model | string | ○ | 使用した Writer モデル名 |

### 2.4 操作仕様

#### 書き込み（append_history）

1. ロック取得
2. ファイル読み込み（存在しない場合は空リスト）
3. 新エントリを先頭 (index 0) に挿入
4. エントリ数 > 50 の場合: 末尾から削除し、対応する PPTX ファイルも `os.remove()` で削除
5. ファイル書き込み（インデント 2 スペース）
6. ロック解放

#### 読み込み（list_history）

- `output_path` フィールドを除いた辞書を返却（セキュリティ）
- ファイルが実際に存在するエントリのみ返却
- 最大 n 件（デフォルト 20 件）

#### 単件取得（get_history_item）

- `job_id` で検索
- `output_path` を含む完全なエントリを返却（ダウンロード用）

---

## 3. ChromaDB — ベクターストア

### 3.1 概要

| 項目 | 内容 |
|-----|------|
| ディレクトリ | `data/chroma_db/` |
| ライブラリ | chromadb ≥ 1.0 |
| ストレージ形式 | PersistentClient（SQLite ベース） |
| コレクション名 | `past_reports` |
| 距離メトリック | cosine（コサイン類似度） |
| 再起動耐性 | あり（永続化） |

### 3.2 コレクション定義

```python
collection = client.get_or_create_collection(
    name="past_reports",
    metadata={"hnsw:space": "cosine"}
)
```

### 3.3 ドキュメントスキーマ

各チャンクは以下の形式で保存される。

| 項目 | 型 | 説明 |
|----|-----|------|
| id | string | `{file_id}_{chunk_idx}` 形式の一意ID |
| document | string | スライドから抽出したテキストチャンク |
| embedding | float[768] | nomic-embed-text による 768 次元ベクター |
| metadata.filename | string | 元の PPTX ファイル名 |
| metadata.file_id | string | ファイル名の MD5 ハッシュ（16 進数） |
| metadata.chunk_idx | int | スライドのインデックス（0 始まり） |

### 3.4 ID 生成規則

```
file_id = md5(filename).hexdigest()
chunk_id = f"{file_id}_{chunk_idx}"
```

例: ファイル名 `report_2024Q1.pptx`、スライド 3
- file_id: `a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4`
- chunk_id: `a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4_2`

### 3.5 CRUD 操作仕様

#### 登録（register_report）

```
入力: pptx_path (str), filename (str)
処理:
  1. python-pptx でスライドごとにテキスト抽出
  2. {{...}} プレースホルダーをタグ除去
  3. 30 文字未満チャンクを除外
  4. 既存の file_id を持つドキュメントを削除（重複登録防止）
  5. 各チャンクを nomic-embed-text でベクター化（POST /api/embeddings）
  6. collection.add(ids, embeddings, documents, metadatas) で一括保存
出力: チャンク数 (int)
```

#### 検索（search_context）

```
入力: query (str), n_results (int) = 5
処理:
  1. query を nomic-embed-text でベクター化
  2. collection.query(query_embeddings, n_results) で類似検索
  3. 結果の distance でフィルタリング
     採用条件: distance < (1 - RAG_THRESHOLD) = 0.7
  4. 採用チャンクを "\n\n---\n\n" で連結
  5. 先頭 1,500 文字に切り詰め
出力: コンテキスト文字列（マッチなし時は空文字列）
```

#### 一覧取得（list_registered）

```
処理:
  1. collection.get(include=["metadatas"]) で全メタデータ取得
  2. file_id でグループ化してユニーク一覧を生成
出力: [{filename, file_id, chunks}]
```

#### チャンク取得（get_chunks_for_file）

```
入力: file_id (str)
処理:
  1. collection.get(where={"file_id": file_id}) で全チャンク取得
  2. chunk_idx でソート
出力: [{id, text, chunk_idx}]
```

#### 削除（delete_report）

```
入力: file_id (str)
処理:
  1. collection.get(where={"file_id": file_id}) で対象チャンク ID を取得
  2. collection.delete(ids) で一括削除
出力: 削除チャンク数 (int)
```

### 3.6 RAG 検索パラメータ

| パラメータ | 値 | 説明 |
|---------|---|------|
| n_results | 5 | 検索時の取得件数 |
| 類似度閾値 | 0.3 | cosine 類似度（distance < 0.7 を採用） |
| 最大コンテキスト文字数 | 1,500 | Writer プロンプトに付加する上限 |
| 最小チャンク文字数 | 30 | 登録時に短すぎるチャンクを除外 |

---

## 4. ファイルシステムストレージ

### 4.1 生成済み PPTX ファイル

| 項目 | 内容 |
|-----|------|
| 保存先 | `output/` ディレクトリ |
| ファイル名規則 | `report_{job_id}.pptx` |
| job_id 形式 | `uuid.uuid4().hex[:12]`（12 文字の16進数） |
| ライフタイム | 最大 50 件保持（超過分を自動削除） |

例: `output/report_ab12cd34ef56.pptx`

### 4.2 テンプレートファイル

| 項目 | 内容 |
|-----|------|
| 保存先 | `data/` ディレクトリ |
| ファイル名規則 | `template*.pptx` |
| 管理 | 手動配置（API による登録なし） |

### 4.3 一時ファイル

| 種別 | 保存先 | 削除タイミング |
|-----|------|------------|
| アップロード Excel/CSV | `tempfile.mkdtemp()` | 生成処理の `finally` ブロック |
| アップロード PPTX テンプレート | `tempfile.mkdtemp()` | 生成処理の `finally` ブロック |
| 参照 PPTX (RAG 登録用) | `tempfile.mkdtemp()` | 登録処理の `finally` ブロック |

---

## 5. Ollama API データフォーマット

Ollama はローカル HTTP サービスとして動作する。以下はシステムが送受信するデータ形式。

### 5.1 テキスト生成リクエスト（POST /api/generate）

```json
{
  "model": "qwen3:8b",
  "prompt": "...",
  "stream": true
}
```

### 5.2 テキスト生成レスポンス（ストリーミング、行区切り JSON）

```json
{"model":"qwen3:8b","created_at":"...","response":"こんに","done":false}
{"model":"qwen3:8b","created_at":"...","response":"ちは","done":false}
{"model":"qwen3:8b","created_at":"...","response":"","done":true}
```

### 5.3 埋め込みリクエスト（POST /api/embeddings）

```json
{
  "model": "nomic-embed-text",
  "prompt": "埋め込みを生成するテキスト"
}
```

### 5.4 埋め込みレスポンス

```json
{
  "embedding": [0.0123, -0.0456, 0.0789, ...]
}
```

- ベクター次元数: 768

### 5.5 モデル一覧（GET /api/tags）

```json
{
  "models": [
    {
      "name": "qwen2.5:3b",
      "modified_at": "2026-03-01T...",
      "size": 1234567890
    },
    {
      "name": "qwen3:8b",
      ...
    }
  ]
}
```

---

## 6. Analyst AI 出力 JSON スキーマ

Analyst AI（qwen2.5:3b）が出力する中間データの仕様。
PPTX には直接使用せず、Writer AI のプロンプト構築に使用する。

```json
{
  "period": "2024年1月〜2024年12月",
  "total_sales": 12345678,
  "total_qty": 5678,
  "top_products": [
    {"name": "商品A", "amount": 3456789},
    {"name": "商品B", "amount": 2345678}
  ],
  "bottom_products": [
    {"name": "商品Z", "amount": 12345}
  ],
  "top_regions": [
    {"name": "東京", "amount": 4567890}
  ],
  "bottom_regions": [
    {"name": "沖縄", "amount": 23456}
  ],
  "top_reps": [
    {"name": "山田太郎", "amount": 2345678}
  ],
  "key_facts": [
    "Q3 売上が全四半期中最高",
    "商品Aが総売上の28%を占める"
  ],
  "concerns": [
    "地方地域の売上が前年比10%減",
    "担当者Bの売上が大幅に低下"
  ],
  "yoy_change": {
    "2023": "+12.5%",
    "2024": "+8.3%"
  }
}
```

### 6.1 フィールド定義

| フィールド名 | 型 | 説明 |
|------------|----|----|
| period | string | 集計対象期間のテキスト |
| total_sales | number | 総売上金額（円） |
| total_qty | number | 総販売数量 |
| top_products | array | 上位商品（名前・金額） |
| bottom_products | array | 下位商品（名前・金額） |
| top_regions | array | 上位地域（名前・金額） |
| bottom_regions | array | 下位地域（名前・金額） |
| top_reps | array | 上位担当者（名前・金額） |
| key_facts | string[] | 注目すべき事実の箇条書き |
| concerns | string[] | 懸念事項の箇条書き |
| yoy_change | object | {年: 前年比パーセント文字列} |

---

## 7. 入力データ仕様（Excel/CSV）

### 7.1 必須列

| 列名 | 型 | 説明 | 例 |
|----|----|----|---|
| 日付 | date/string | 売上日 | 2024-01-15 |
| 商品名 | string | 商品名 | 商品A |
| 担当者 | string | 営業担当者名 | 山田太郎 |
| 地域 | string | 販売地域名 | 東京 |
| 数量 | integer | 販売数量 | 10 |
| 売上金額 | integer | 売上金額（円） | 100000 |

### 7.2 任意列

| 列名 | 型 | 説明 | 効果 |
|----|----|----|------|
| 利益額 | integer | 利益金額（円） | グラフに利益率折れ線を追加 |

### 7.3 対応ファイル形式

| 形式 | 拡張子 | 文字コード |
|-----|--------|---------|
| Excel | .xlsx | UTF-8（openpyxl） |
| Excel（旧形式） | .xls | UTF-8（openpyxl） |
| CSV | .csv | UTF-8 → Shift-JIS の順でリトライ |

---

## 8. ディレクトリ・ファイル一覧

```
makeReportOllama/
├── output/
│   ├── history.json              # 生成履歴 JSON
│   └── report_{job_id}.pptx     # 生成済み PPTX（最大 50 件）
│
├── data/
│   ├── template.pptx             # デフォルトテンプレート
│   ├── template_consultant.pptx  # コンサルテンプレート
│   ├── template_executive.pptx   # エグゼクティブテンプレート
│   ├── sample_*.xlsx/.csv        # サンプルデータファイル
│   └── chroma_db/                # ChromaDB 永続化ディレクトリ
│       ├── chroma.sqlite3        # ChromaDB メタデータ
│       └── ...（ChromaDB 内部ファイル）
│
└── backend/
    └── app.log                   # アプリケーションログ（追記）
```
