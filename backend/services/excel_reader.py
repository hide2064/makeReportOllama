"""
excel_reader.py
Excel ファイルを読み込み、売上データを集計して要約テキストを返す。
"""

import logging
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)


def read_and_summarize(excel_path: str) -> dict:
    """
    Excel を読み込み、Ollama に渡すための集計サマリーを返す。

    Returns:
        {
          "total_amount": int,
          "total_qty": int,
          "by_product": str,
          "by_region": str,
          "by_rep": str,
          "period": str,
          "raw_summary": str,  # Ollama プロンプトに埋め込む文字列
        }
    """
    logger.info(f"Excel 読み込み開始: {excel_path}")
    if not Path(excel_path).exists():
        raise FileNotFoundError(f"Excel ファイルが見つかりません: {excel_path}")

    df = pd.read_excel(excel_path)
    logger.info(f"読み込み行数: {len(df)}")

    required_cols = {"日付", "商品名", "担当者", "地域", "数量", "売上金額"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"必須列が不足しています: {missing}")

    df["日付"] = pd.to_datetime(df["日付"])

    total_amount = int(df["売上金額"].sum())
    total_qty    = int(df["数量"].sum())
    period_start = df["日付"].min().strftime("%Y/%m/%d")
    period_end   = df["日付"].max().strftime("%Y/%m/%d")

    by_product = (
        df.groupby("商品名")["売上金額"]
        .sum()
        .sort_values(ascending=False)
        .apply(lambda x: f"{x:,}円")
        .to_string()
    )
    by_region = (
        df.groupby("地域")["売上金額"]
        .sum()
        .sort_values(ascending=False)
        .apply(lambda x: f"{x:,}円")
        .to_string()
    )
    by_rep = (
        df.groupby("担当者")["売上金額"]
        .sum()
        .sort_values(ascending=False)
        .apply(lambda x: f"{x:,}円")
        .to_string()
    )

    raw_summary = (
        f"集計期間: {period_start} ～ {period_end}\n"
        f"総売上金額: {total_amount:,}円\n"
        f"総販売数量: {total_qty}個\n\n"
        f"【商品別売上】\n{by_product}\n\n"
        f"【地域別売上】\n{by_region}\n\n"
        f"【担当者別売上】\n{by_rep}\n"
    )

    logger.info("集計完了")
    return {
        "total_amount": total_amount,
        "total_qty":    total_qty,
        "by_product":   by_product,
        "by_region":    by_region,
        "by_rep":       by_rep,
        "period":       f"{period_start} ～ {period_end}",
        "raw_summary":  raw_summary,
    }
