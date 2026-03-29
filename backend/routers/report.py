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
import uuid
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from config import ANALYST_MAX_RETRIES, DATA_DIR, MODEL_ANALYST, MODEL_WRITER, OUTPUT_DIR
from services.excel_reader import read_and_summarize
from services.history_store import append_history
from services.ollama_client import (
    build_analyst_prompt,
    build_writer_prompt,
    check_ollama,
    generate,
    parse_analyst_json,
    parse_writer_response,
)
from services.pptx_generator import generate_pptx
from services.rag_store import search_context

logger = logging.getLogger(__name__)
router = APIRouter()

OUTPUT_DIR.mkdir(exist_ok=True)

# ── 処理状態（スレッド間共有） ────────────────────────────────
_status_lock = threading.Lock()
_status: dict = {"step": "", "done": False, "error": "", "output_path": ""}
_executor = ThreadPoolExecutor(max_workers=1)


def _set_status(**kwargs):
    with _status_lock:
        _status.update(kwargs)
    logger.info(f"Status: {kwargs}")


# ── バックグラウンド処理（同期・スレッド実行） ─────────────────
def _run_generation(
    excel_data: bytes,
    template_data: Optional[bytes],
    excel_filename: str,
    template_name: str = "",
    slide_options: Optional[dict] = None,
    analyst_model: str = "",
    writer_model: str = "",
    date_from: str = "",
    date_to: str = "",
    extra_context: str = "",
):
    tmpdir = tempfile.mkdtemp()
    try:
        suffix      = Path(excel_filename).suffix.lower() or ".xlsx"
        excel_path  = os.path.join(tmpdir, f"sales_data{suffix}")
        # UUID ベースの出力パス（同時アクセス時の上書き競合を防ぐ）
        job_id      = uuid.uuid4().hex[:12]
        output_path = str(OUTPUT_DIR / f"report_{job_id}.pptx")
        # モデル名の解決（空文字はデフォルト値を使用）
        used_analyst = analyst_model or MODEL_ANALYST
        used_writer  = writer_model  or MODEL_WRITER

        with open(excel_path, "wb") as f:
            f.write(excel_data)

        # テンプレートパスの解決
        if template_name:
            template_path = str(DATA_DIR / template_name)
            if not Path(template_path).exists():
                _set_status(error=f"テンプレートが見つかりません: {template_name}")
                return
        elif template_data:
            template_path = os.path.join(tmpdir, "template.pptx")
            with open(template_path, "wb") as f:
                f.write(template_data)
        else:
            _set_status(error="テンプレートが指定されていません。ファイルをアップロードするか、サーバーのテンプレートを選択してください。")
            return

        # Step 0: Ollama プリフライトチェック（接続 + モデル存在確認）
        _set_status(step="[0/3]  Ollama の接続とモデルを確認しています...")
        try:
            check_ollama(used_analyst)
            check_ollama(used_writer)
        except RuntimeError as e:
            _set_status(error=str(e))
            return

        # Step 1
        _set_status(step="[1/3]  Excel / CSV を読み込んでいます...")
        try:
            summary_data = read_and_summarize(excel_path, date_from=date_from, date_to=date_to)
        except (FileNotFoundError, ValueError) as e:
            _set_status(error=str(e))
            return

        # Step 2a: Analyst AI — 数値データを構造化 JSON に変換（最大 3 回試行）
        _set_status(step=f"[2/3]  売上データを解析中です... (Analyst: {used_analyst})\n"
                         "       数値・トレンドを抽出しています。")

        def _on_analyst_token(count: int):
            _set_status(step=f"[2/3]  売上データを解析中です... (Analyst: {used_analyst})\n"
                             f"       生成中... {count} トークン生成済み")

        analyst_data: dict = {}
        analyst_prompt = build_analyst_prompt(summary_data["raw_summary"])
        for attempt in range(1, ANALYST_MAX_RETRIES + 1):
            try:
                analyst_raw  = generate(analyst_prompt, model=used_analyst,
                                        on_token=_on_analyst_token)
                analyst_data = parse_analyst_json(analyst_raw)
            except RuntimeError as e:
                _set_status(error=str(e))
                return

            if analyst_data:
                logger.info(f"Analyst 結果 (試行 {attempt}): {list(analyst_data.keys())}")
                break

            # JSON が空 → リトライ
            if attempt < ANALYST_MAX_RETRIES:
                logger.warning(f"Analyst JSON が空 (試行 {attempt}/{ANALYST_MAX_RETRIES})。リトライします。")
                _set_status(step=f"[2/3]  売上データを再解析中です... (試行 {attempt + 1}/{ANALYST_MAX_RETRIES})")
            else:
                # 最大試行回数を超えた場合は raw_summary でフォールバック
                logger.warning("Analyst JSON の取得に失敗。raw_summary でフォールバックします。")
                analyst_data = {}

        # Step 2b: RAG — 過去レポートから類似コンテキストを取得
        _set_status(step="[2/3]  過去レポートから関連情報を検索中です... (RAG)")
        rag_context = search_context(summary_data["raw_summary"])
        if rag_context:
            logger.info(f"RAG コンテキスト取得: {len(rag_context)} 字")
        else:
            logger.info("RAG コンテキストなし（過去資料未登録 or 類似度低）")

        # Step 2c: Writer AI — 構造化データ + RAG 文脈から日本語ビジネス文章を生成
        _set_status(step=f"[2/3]  レポート文章を生成中です... (Writer: {used_writer})\n"
                         "       CPU 推論のため数分かかります。このまましばらくお待ちください。")

        def _on_writer_token(count: int):
            _set_status(step=f"[2/3]  レポート文章を生成中です... (Writer: {used_writer})\n"
                             f"       生成中... {count} トークン生成済み")

        try:
            writer_raw = generate(
                build_writer_prompt(
                    analyst_data, summary_data["raw_summary"],
                    rag_context, extra_context,
                ),
                model=used_writer,
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
                quarterly_region_pivot=summary_data.get("quarterly_region_pivot"),
                quarterly_rep_pivot=summary_data.get("quarterly_rep_pivot"),
                monthly_margin=summary_data.get("monthly_margin"),
                slide_options=slide_options,
            )
        except Exception as e:
            _set_status(error=f"PPTX 生成に失敗しました: {e}")
            return

        # H-4: 履歴に追記
        append_history(
            job_id=job_id,
            original_filename=excel_filename,
            output_path=output_path,
            analyst_model=used_analyst,
            writer_model=used_writer,
        )
        _set_status(step="完了しました！", done=True, output_path=output_path)

    except Exception as e:
        logger.exception("予期しないエラー")
        _set_status(error=str(e))
    finally:
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)


# ── エンドポイント ─────────────────────────────────────────────

@router.post("/api/generate")
async def generate_report(
    excel_file:          UploadFile      = File(..., description="売上データ Excel/CSV"),
    template_file:       Optional[UploadFile] = File(None, description="PPTX テンプレート (.pptx)"),
    template_name:       str             = Form("", description="サーバー上のテンプレート名"),
    slide_product_table: bool            = Form(True),
    slide_region_table:  bool            = Form(False),
    slide_rep_table:     bool            = Form(False),
    slide_chart:         bool            = Form(True),
    chart_product_type:  str             = Form("bar"),
    analyst_model:       str             = Form(""),
    writer_model:        str             = Form(""),
    date_from:           str             = Form(""),
    date_to:             str             = Form(""),
    extra_context:       str             = Form("", description="追加プロンプト（オプション）"),
):
    """処理をバックグラウンドスレッドで開始し、即座に返す。"""
    # 処理中の場合は 409 を返す（二重投入防止）
    with _status_lock:
        busy = _status["step"] and not _status["done"] and not _status["error"]
    if busy:
        raise HTTPException(status_code=409, detail="現在レポートを生成中です。処理が完了してから再度お試しください。")

    # 前回ジョブの出力ファイルを削除（ディスク節約）
    with _status_lock:
        prev_path = _status.get("output_path", "")
    if prev_path and Path(prev_path).exists():
        try:
            Path(prev_path).unlink()
            logger.info(f"前回出力ファイルを削除: {prev_path}")
        except OSError:
            pass

    logger.info(f"generate_report 開始: excel={excel_file.filename}, template_name={template_name!r}")
    excel_data    = await excel_file.read()
    template_data = await template_file.read() if template_file else None
    excel_filename = excel_file.filename or "sales_data.xlsx"

    slide_options = {
        "product_table":       slide_product_table,
        "region_table":        slide_region_table,
        "rep_table":           slide_rep_table,
        "chart":               slide_chart,
        "chart_product_type":  chart_product_type,
    }

    _set_status(step="[0/3]  Ollama の接続とモデルを確認しています...", done=False, error="", output_path="")
    _executor.submit(
        _run_generation,
        excel_data, template_data, excel_filename,
        template_name, slide_options,
        analyst_model, writer_model,
        date_from, date_to,
        extra_context,
    )

    return {"status": "started"}


@router.get("/api/progress")
async def get_progress():
    """現在の処理ステップを返す。フロントエンドがポーリングで呼び出す。"""
    with _status_lock:
        return dict(_status)


@router.get("/api/download")
async def download_report():
    """生成済み PPTX をダウンロードさせる。"""
    with _status_lock:
        output_path = _status.get("output_path", "")
    if not output_path or not Path(output_path).exists():
        raise HTTPException(status_code=404, detail="レポートファイルが見つかりません。先に生成してください。")
    return FileResponse(
        path=output_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="report.pptx",
    )
