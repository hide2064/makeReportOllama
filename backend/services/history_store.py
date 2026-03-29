"""
services/history_store.py
生成済みレポートのメタデータを output/history.json に保存・取得する。
スレッドセーフ設計（_lock で保護）。
"""

import json
import logging
import threading
from datetime import datetime, timezone
from pathlib import Path

logger = logging.getLogger(__name__)

OUTPUT_DIR   = Path(__file__).parent.parent.parent / "output"
HISTORY_FILE = OUTPUT_DIR / "history.json"
MAX_ENTRIES  = 50

_lock: threading.Lock = threading.Lock()


def _load() -> list[dict]:
    if not HISTORY_FILE.exists():
        return []
    try:
        return json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
    except Exception as e:
        logger.warning(f"history.json 読み込み失敗: {e}")
        return []


def _save(entries: list[dict]) -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    HISTORY_FILE.write_text(
        json.dumps(entries, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def append_history(
    job_id: str,
    original_filename: str,
    output_path: str,
    analyst_model: str = "",
    writer_model: str = "",
) -> None:
    """生成完了したレポートのメタデータを追記する（最大 MAX_ENTRIES 件でローテーション）。"""
    entry = {
        "job_id":            job_id,
        "created_at":        datetime.now(timezone.utc).isoformat(),
        "original_filename": original_filename,
        "output_path":       output_path,
        "analyst_model":     analyst_model,
        "writer_model":      writer_model,
    }
    with _lock:
        entries = _load()
        entries.insert(0, entry)          # 先頭に追加（新しい順）
        # ローテーション: 溢れた分の PPTX ファイルも削除
        evicted = entries[MAX_ENTRIES:]
        entries = entries[:MAX_ENTRIES]
        _save(entries)
    for old in evicted:
        old_path = old.get("output_path", "")
        if old_path and Path(old_path).exists():
            try:
                Path(old_path).unlink()
                logger.info(f"古いレポートを削除: {old_path}")
            except OSError as e:
                logger.warning(f"古いレポート削除失敗: {e}")
    logger.info(f"履歴追記: job_id={job_id}")


def list_history(n: int = 20) -> list[dict]:
    """直近 n 件を返す。output_path は除外してセキュリティを確保する。"""
    with _lock:
        entries = _load()
    safe = []
    for e in entries[:n]:
        # ファイルが実際に存在するものだけ返す
        if Path(e.get("output_path", "")).exists():
            safe.append({
                "job_id":            e["job_id"],
                "created_at":        e["created_at"],
                "original_filename": e.get("original_filename", ""),
                "analyst_model":     e.get("analyst_model", ""),
                "writer_model":      e.get("writer_model", ""),
            })
    return safe


def get_history_item(job_id: str) -> dict | None:
    """job_id でエントリを検索し output_path を含む生データを返す。"""
    with _lock:
        entries = _load()
    for e in entries:
        if e.get("job_id") == job_id:
            return e
    return None
