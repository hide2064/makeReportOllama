"""
test_excel_reader.py — excel_reader サービスのユニットテスト
"""
import os
import tempfile

import pandas as pd
import pytest

from services.excel_reader import read_and_summarize


@pytest.fixture()
def sample_excel(tmp_path):
    """テスト用の最小限 Excel ファイルを生成して返す。"""
    data = {
        "日付":     ["2025-01-10", "2025-01-20", "2025-02-05"],
        "商品名":   ["商品A", "商品B", "商品A"],
        "担当者":   ["田中", "佐藤", "田中"],
        "地域":     ["東京", "大阪", "東京"],
        "数量":     [2, 1, 3],
        "売上金額": [100_000, 80_000, 150_000],
    }
    df = pd.DataFrame(data)
    path = str(tmp_path / "test_sales.xlsx")
    df.to_excel(path, index=False)
    return path


def test_returns_required_keys(sample_excel):
    result = read_and_summarize(sample_excel)
    for key in ("total_amount", "total_qty", "period", "raw_summary",
                "by_product", "by_region", "by_rep"):
        assert key in result, f"key '{key}' が結果に含まれていません"


def test_total_amount(sample_excel):
    result = read_and_summarize(sample_excel)
    assert result["total_amount"] == 330_000


def test_total_qty(sample_excel):
    result = read_and_summarize(sample_excel)
    assert result["total_qty"] == 6


def test_file_not_found():
    with pytest.raises(FileNotFoundError):
        read_and_summarize("/nonexistent/path/sales.xlsx")


def test_missing_column(tmp_path):
    df = pd.DataFrame({"日付": ["2025-01-01"], "商品名": ["商品A"]})
    path = str(tmp_path / "bad.xlsx")
    df.to_excel(path, index=False)
    with pytest.raises(ValueError, match="必須列"):
        read_and_summarize(path)
