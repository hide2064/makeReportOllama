"""
config.py
アプリケーション全体で共有する定数。
環境変数で上書き可能なものは os.environ.get() で取得する。
"""

import os
from pathlib import Path

# ── ディレクトリ ────────────────────────────────────────────────
ROOT_DIR   = Path(__file__).parent.parent
DATA_DIR   = ROOT_DIR / "data"
OUTPUT_DIR = ROOT_DIR / "output"

# ── Ollama ─────────────────────────────────────────────────────
OLLAMA_BASE_URL   = os.environ.get("OLLAMA_BASE_URL", "http://localhost:11434")
OLLAMA_GENERATE   = f"{OLLAMA_BASE_URL}/api/generate"
OLLAMA_EMBED_URL  = f"{OLLAMA_BASE_URL}/api/embeddings"
OLLAMA_TAGS_URL   = f"{OLLAMA_BASE_URL}/api/tags"

# モデル名（環境変数で上書き可）
MODEL_ANALYST = os.environ.get("OLLAMA_MODEL_ANALYST", "qwen2.5:3b")
MODEL_WRITER  = os.environ.get("OLLAMA_MODEL_WRITER",  "qwen3:8b")
MODEL_EMBED   = os.environ.get("OLLAMA_MODEL_EMBED",   "nomic-embed-text")

# タイムアウト (秒)
OLLAMA_CONNECT_TIMEOUT  = 10
OLLAMA_GENERATE_TIMEOUT = int(os.environ.get("OLLAMA_TIMEOUT", "1200"))  # 20分
OLLAMA_EMBED_TIMEOUT    = 60

# ── RAG ────────────────────────────────────────────────────────
CHROMA_DIR       = DATA_DIR / "chroma_db"
RAG_COLLECTION   = "past_reports"
RAG_MAX_CHUNKS   = 5
RAG_MAX_CTX_CHARS= 1_500
RAG_MIN_CHUNK    = 30
RAG_SIM_THRESHOLD= 0.3

# ── 履歴 ───────────────────────────────────────────────────────
HISTORY_FILE    = OUTPUT_DIR / "history.json"
HISTORY_MAX     = 50

# ── 生成パイプライン ────────────────────────────────────────────
ANALYST_MAX_RETRIES = 3
