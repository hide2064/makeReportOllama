# 次フェーズ作業候補（バックログ）

最終更新: 2026-03-29

---

## 優先度高（未着手）

_現時点でなし_

---

## 優先度中（未着手）

_現時点でなし_

---

## 優先度低（未着手）

_現時点でなし_

---

## 実装済み（参考）

- Phase 1: Excel 読み込み → PPTX テンプレート置換
- Phase 2: 2 モデルパイプライン（Analyst: qwen2.5:3b / Writer: qwen3:8b）
- Phase 3: RAG（ChromaDB + nomic-embed-text）+ 過去資料管理 UI
- Phase 4a: 商品別四半期売上表スライド（python-pptx テーブル）
- Phase 4b: 月次推移バーチャート + 商品別横棒/円グラフスライド（matplotlib）
- 高優先度改善: UUID 出力ファイル管理、Analyst リトライ（最大3回）、CSV サポート、利益データ列対応
- 中優先度改善: GET /api/templates、スライド構成チェックボックス、グラフ種別ラジオ、地域/担当者別表スライド、YoY 前年同期比、汎用 `_add_table_slide`
- 低優先度改善: .gitignore 整備、LOG_LEVEL 環境変数化、pptx_generator ユニットテスト拡充
- H-1: ドラッグ＆ドロップファイルアップロード
- H-2: アップロードデータプレビュー
- H-3: Ollama モデル選択 UI
- H-4: 生成済みレポート履歴
- H-5: 分析期間フィルター
- M-1: ブラウザ内グラフ表示
- M-2: SSE によるリアルタイムストリーミング
- M-3: 列名マッピング UI
- M-4: カスタムコンテキスト入力欄
- M-5: 目標値（KPI）入力と達成率比較
- M-6: PDF 出力オプション
- M-7: RAG コンテキストデバッグパネル
- L-1: ダークモード
- L-2: Docker コンテナ化
- L-3: ブラウザ通知
- L-4: E2E テスト（Playwright）
- L-5: GitHub Actions CI
