"""
main.py
FastAPI アプリケーションのエントリポイント。
"""

import logging
import os
import sys
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

# ── ロギング設定 ─────────────────────────────────────────
LOG_FILE  = Path(__file__).parent / "app.log"
LOG_LEVEL = getattr(logging, os.environ.get("LOG_LEVEL", "INFO").upper(), logging.INFO)
logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)

# ── FastAPI アプリ ────────────────────────────────────────
app = FastAPI(
    title="makeReportOllama API",
    description="Excel 売上データから Ollama を使って PowerPoint 報告書を自動生成する API",
    version="1.0.0",
)

# フロントエンド (Vite dev server) からのアクセスを許可
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── ルーター登録 ─────────────────────────────────────────
from routers.report import router as report_router              # noqa: E402
from routers.references import router as references_router      # noqa: E402
from routers.templates import router as templates_router        # noqa: E402
from routers.models import router as models_router              # noqa: E402
from routers.preview import router as preview_router            # noqa: E402
from routers.history import router as history_router            # noqa: E402

app.include_router(report_router)
app.include_router(references_router)
app.include_router(templates_router)
app.include_router(models_router)
app.include_router(preview_router)
app.include_router(history_router)


@app.get("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    import uvicorn
    logger.info("FastAPI サーバー起動")
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
