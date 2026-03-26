# makeReportOllama — 開発メモ

## 次フェーズ作業：Phase 3 RAG 実装

### 概要
過去に生成した複数の報告書 PPTX を参照情報として活用し、
新規レポート生成時の文章品質と文脈の一貫性を向上させる。

### 追加コンポーネント
- **埋め込みモデル**: `nomic-embed-text`（Ollama, 274MB, CPU対応）
- **ベクターDB**: `ChromaDB`（Python ライブラリ、サーバー不要）
- **文書読込**: `python-pptx`（既存）

### 処理フロー
```
【事前処理 — 過去資料登録時】
  過去 PPTX → テキスト抽出 (python-pptx)
            → チャンク分割 (スライド単位)
            → nomic-embed-text でベクター化
            → ChromaDB に永続保存

【レポート生成時】
  新規 CSV → Analyst AI (qwen2.5:3b) → 構造化 JSON
  + ChromaDB 類似検索 (上位 3〜5 件、各 300 字以内)
  → Writer AI (qwen3:8b) に JSON + 過去文脈を渡す
  → PPTX 生成
```

### 追加インストール
```bash
ollama pull nomic-embed-text
pip install chromadb
```

### 実装上の注意点
- コンテキスト長超過を防ぐため、取得チャンクは合計 1,500 字以内に制限する
- 過去資料が 3 件未満の場合は RAG をスキップして通常生成にフォールバック
- ChromaDB の永続化ディレクトリ: `data/chroma_db/`
- 過去資料の登録 UI は frontend に「過去資料管理」セクションとして追加する

### 参考
- [前回の設計検討メモ] 2モデル分離（Phase 2）が完了してから着手すること
