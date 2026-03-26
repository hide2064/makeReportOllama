"""
test_api.py — FastAPI エンドポイントの結合テスト
Ollama 通信部分は Mock 化してテストする。

Phase 2 対応: generate が Analyst (1回目) と Writer (2回目) で2回呼ばれる。
"""
import io
import os
import time
from unittest.mock import patch

import pandas as pd
import pytest
from fastapi.testclient import TestClient

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../../backend"))

from main import app  # noqa: E402

client = TestClient(app, raise_server_exceptions=True)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "../../data/template.pptx")

# Analyst AI が返す JSON モック
MOCK_ANALYST_JSON = '{"period": "2025/01/10 ~ 2025/02/15", "total_sales": 180000, "total_qty": 3, "top_products": [{"name": "商品A", "amount": 100000}], "bottom_products": [{"name": "商品B", "amount": 80000}], "top_regions": [{"name": "東京", "amount": 100000}], "bottom_regions": [{"name": "大阪", "amount": 80000}], "top_reps": [{"name": "田中", "amount": 100000}], "key_facts": ["売上合計18万円"], "concerns": ["大阪が低迷"]}'

# Writer AI が返すテキストモック
MOCK_WRITER_TEXT = "---SUMMARY---\nMock サマリー\n---ANALYSIS---\nMock 分析"


def _make_excel_bytes() -> bytes:
    data = {
        "日付":     ["2025-01-10", "2025-02-15"],
        "商品名":   ["商品A", "商品B"],
        "担当者":   ["田中", "佐藤"],
        "地域":     ["東京", "大阪"],
        "数量":     [2, 1],
        "売上金額": [100_000, 80_000],
    }
    buf = io.BytesIO()
    pd.DataFrame(data).to_excel(buf, index=False)
    return buf.getvalue()


def _make_csv_bytes() -> bytes:
    data = {
        "日付":     ["2025-01-10", "2025-02-15"],
        "商品名":   ["商品A", "商品B"],
        "担当者":   ["田中", "佐藤"],
        "地域":     ["東京", "大阪"],
        "数量":     [2, 1],
        "売上金額": [100_000, 80_000],
    }
    return pd.DataFrame(data).to_csv(index=False).encode("utf-8")


def _make_pptx_bytes() -> bytes:
    with open(TEMPLATE_PATH, "rb") as f:
        return f.read()


def _post_generate(excel_bytes, filename="sales.xlsx"):
    return client.post(
        "/api/generate",
        files={
            "excel_file":    (filename, excel_bytes, "application/octet-stream"),
            "template_file": ("template.pptx", _make_pptx_bytes(), "application/octet-stream"),
        },
    )


def _wait_done(timeout=10.0):
    """progress が done になるまでポーリング（テスト用）。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        res = client.get("/api/progress")
        data = res.json()
        if data.get("error"):
            return data
        if data.get("done"):
            return data
        time.sleep(0.2)
    return client.get("/api/progress").json()


# ── テスト ──────────────────────────────────────────────────

def test_health_endpoint():
    res = client.get("/health")
    assert res.status_code == 200
    assert res.json() == {"status": "ok"}


@patch("routers.report.generate", side_effect=[MOCK_ANALYST_JSON, MOCK_WRITER_TEXT])
def test_generate_success_excel(mock_generate):
    """正常系(Excel): Analyst + Writer の2回呼び出しで done になる。"""
    res = _post_generate(_make_excel_bytes(), "sales.xlsx")
    assert res.status_code == 200
    assert res.json()["status"] == "started"

    result = _wait_done()
    assert result.get("done") is True
    assert mock_generate.call_count == 2  # Analyst + Writer


@patch("routers.report.generate", side_effect=[MOCK_ANALYST_JSON, MOCK_WRITER_TEXT])
def test_generate_success_csv(mock_generate):
    """正常系(CSV): CSV ファイルでも Analyst + Writer 2回呼び出しで done になる。"""
    res = _post_generate(_make_csv_bytes(), "sales.csv")
    assert res.status_code == 200
    result = _wait_done()
    assert result.get("done") is True
    assert mock_generate.call_count == 2


@patch("routers.report.generate", side_effect=RuntimeError("Ollama に接続できません"))
def test_generate_ollama_error(mock_generate):
    """Analyst 呼び出しでエラー発生時に progress に error が入る。"""
    _post_generate(_make_excel_bytes())
    result = _wait_done()
    assert result.get("error"), f"error expected, got: {result}"


@patch("routers.report.generate", side_effect=[MOCK_ANALYST_JSON, RuntimeError("Writer エラー")])
def test_generate_writer_error(mock_generate):
    """Writer 呼び出しでエラー発生時に progress に error が入る。"""
    _post_generate(_make_excel_bytes())
    result = _wait_done()
    assert result.get("error"), f"error expected, got: {result}"


def test_generate_invalid_excel():
    """不正なファイルを送ると progress に error が入る。"""
    _post_generate(b"not an excel", "bad.xlsx")
    result = _wait_done()
    assert result.get("error"), f"error expected, got: {result}"


def test_download_not_found():
    """レポート未生成時に 404 が返ることを確認。"""
    import routers.report as report_module
    # output_path を空にリセットして「未生成」状態を作る
    with report_module._status_lock:
        prev = report_module._status.get("output_path", "")
        report_module._status["output_path"] = ""
    try:
        res = client.get("/api/download")
        assert res.status_code == 404
    finally:
        with report_module._status_lock:
            report_module._status["output_path"] = prev
