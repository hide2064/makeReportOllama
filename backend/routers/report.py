"""
routers/report.py
POST /api/generate  — バックグラウンドスレッドで処理を開始し即返却。
GET  /api/progress  — 処理ステップをポーリングで取得。
GET  /api/download  — 生成済み PPTX をダウンロード。
"""

import logging
import os
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from services.excel_reader import read_and_summarize
from services.rag_store import search_context
from services.ollama_client import (
    MODEL_ANALYST,
    MODEL_WRITER,
    build_analyst_prompt,
    build_writer_prompt,
    generate,
    parse_analyst_json,
    parse_writer_response,
)
from services.pptx_generator import generate_pptx

logger = logging.getLogger(__name__)
router = APIRouter()

OUTPUT_DIR = Path(__file__).parent.parent.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# ── 処理状態（スレッド間共有） ────────────────────────────────
_status_lock = threading.Lock()
_status: dict = {"step": "", "done": False, "error": ""}
_executor = ThreadPoolExecutor(max_workers=1)


def _set_status(**kwargs):
    with _status_lock:
        _status.update(kwargs)
    logger.info(f"Status: {kwargs}")


# ── バックグラウンド処理（同期・スレッド実行） ─────────────────
def _run_generation(excel_data: bytes, template_data: bytes, excel_filename: str):
    tmpdir = tempfile.mkdtemp()
    try:
        suffix        = Path(excel_filename).suffix.lower() or ".xlsx"
        excel_path    = os.path.join(tmpdir, f"sales_data{suffix}")
        template_path = os.path.join(tmpdir, "template.pptx")
        output_path   = str(OUTPUT_DIR / "report.pptx")

        with open(excel_path, "wb") as f:
            f.write(excel_data)
        with open(template_path, "wb") as f:
            f.write(template_data)

        # Step 1
        _set_status(step="[1/3]  Excel / CSV を読み込んでいます...")
        try:
            summary_data = read_and_summarize(excel_path)
        except (FileNotFoundError, ValueError) as e:
            _set_status(error=str(e))
            return

        # Step 2a: Analyst AI — 数値データを構造化 JSON に変換
        _set_status(step=f"[2/3]  売上データを解析中です... (Analyst: {MODEL_ANALYST})\n"
                         "       数値・トレンドを抽出しています。")

        def _on_analyst_token(count: int):
            _set_status(step=f"[2/3]  売上データを解析中です... (Analyst: {MODEL_ANALYST})\n"
                             f"       生成中... {count} トークン生成済み")

        try:
            analyst_raw = generate(
                build_analyst_prompt(summary_data["raw_summary"]),
                model=MODEL_ANALYST,
                on_token=_on_analyst_token,
            )
            analyst_data = parse_analyst_json(analyst_raw)
            logger.info(f"Analyst 結果: {list(analyst_data.keys())}")
        except RuntimeError as e:
            _set_status(error=str(e))
            return

        # Step 2b: RAG — 過去レポートから類似コンテキストを取得
        _set_status(step="[2/3]  過去レポートから関連情報を検索中です... (RAG)")
        rag_context = search_context(summary_data["raw_summary"])
        if rag_context:
            logger.info(f"RAG コンテキスト取得: {len(rag_context)} 字")
        else:
            logger.info("RAG コンテキストなし（過去資料未登録 or 類似度低）")

        # Step 2c: Writer AI — 構造化データ + RAG 文脈から日本語ビジネス文章を生成
        _set_status(step=f"[2/3]  レポート文章を生成中です... (Writer: {MODEL_WRITER})\n"
                         "       CPU 推論のため数分かかります。このまましばらくお待ちください。")

        def _on_writer_token(count: int):
            _set_status(step=f"[2/3]  レポート文章を生成中です... (Writer: {MODEL_WRITER})\n"
                             f"       生成中... {count} トークン生成済み")

        try:
            writer_raw = generate(
                build_writer_prompt(analyst_data, summary_data["raw_summary"], rag_context),
                model=MODEL_WRITER,
                on_token=_on_writer_token,
            )
            summary_text, analysis_text = parse_writer_response(writer_raw)
        except RuntimeError as e:
            _set_status(error=str(e))
            return

        # Step 3
        _set_status(step="[3/3]  PowerPoint レポートを生成しています...")
        try:
            generate_pptx(
                template_path=template_path,
                output_path=output_path,
                summary_text=summary_text,
                analysis_text=analysis_text,
                period=summary_data["period"],
                monthly_totals=summary_data.get("monthly_totals"),
                product_totals=summary_data.get("product_totals"),
                quarterly_product_pivot=summary_data.get("quarterly_product_pivot"),
                monthly_margin=summary_data.get("monthly_margin"),
            )
        except Exception as e:
            _set_status(error=f"PPTX 生成に失敗しました: {e}")
            return

        _set_status(step="完了しました！", done=True)

    except Exception as e:
        logger.exception("予期しないエラー")
        _set_status(error=str(e))
    finally:
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)


# ── エンドポイント ─────────────────────────────────────────────

@router.post("/api/generate")
async def generate_report(
    excel_file:    UploadFile = File(..., description="売上データ Excel/CSV"),
    template_file: UploadFile = File(..., description="PPTX テンプレート (.pptx)"),
):
    """処理をバックグラウンドスレッドで開始し、即座に返す。"""
    # 処理中の場合は 409 を返す（二重投入防止）
    with _status_lock:
        busy = _status["step"] and not _status["done"] and not _status["error"]
    if busy:
        raise HTTPException(status_code=409, detail="現在レポートを生成中です。処理が完了してから再度お試しください。")

    logger.info(f"generate_report 開始: excel={excel_file.filename}")
    excel_data    = await excel_file.read()
    template_data = await template_file.read()
    excel_filename = excel_file.filename or "sales_data.xlsx"

    _set_status(step="[1/3]  Excel / CSV を読み込んでいます...", done=False, error="")
    _executor.submit(_run_generation, excel_data, template_data, excel_filename)

    return {"status": "started"}


@router.get("/api/progress")
async def get_progress():
    """現在の処理ステップを返す。フロントエンドがポーリングで呼び出す。"""
    with _status_lock:
        return dict(_status)


@router.get("/api/download")
async def download_report():
    """生成済み report.pptx をダウンロードさせる。"""
    output_path = OUTPUT_DIR / "report.pptx"
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="レポートファイルが見つかりません。先に生成してください。")
    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="report.pptx",
    )
