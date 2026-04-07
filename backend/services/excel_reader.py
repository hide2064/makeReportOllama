"""
excel_reader.py
Excel / CSV ファイルを読み込み、売上データを集計して要約テキストを返す。
"""

import logging
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)


def read_and_summarize(file_path: str, date_from: str = "", date_to: str = "") -> dict:
    """
    Excel (.xlsx) または CSV (.csv) を読み込み、Ollama に渡すための集計サマリーを返す。

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
    logger.info(f"ファイル読み込み開始: {file_path}")
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {file_path}")

    suffix = path.suffix.lower()
    if suffix == ".csv":
        # UTF-8 → Shift-JIS の順でフォールバック
        try:
            df = pd.read_csv(file_path, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, encoding="shift-jis")
        logger.info("CSV として読み込み")
    elif suffix in (".xlsx", ".xls"):
        df = pd.read_excel(file_path)
        logger.info("Excel として読み込み")
    else:
        raise ValueError(f"対応していないファイル形式です: {suffix}  (.xlsx / .csv のみ対応)")

    logger.info(f"読み込み行数: {len(df)}")

    required_cols = {"日付", "商品名", "担当者", "地域", "数量", "売上金額"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"必須列が不足しています: {missing}  (必須: {sorted(required_cols)})")

    df["日付"] = pd.to_datetime(df["日付"])

    # 分析期間フィルター
    if date_from:
        df = df[df["日付"] >= pd.Timestamp(date_from)]
    if date_to:
        df = df[df["日付"] <= pd.Timestamp(date_to)]
    if df.empty:
        raise ValueError("指定された期間にデータが存在しません。日付範囲を確認してください。")

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

    # 事業部・課 別集計（列が存在する場合のみ）
    by_division = ""
    if "事業部" in df.columns:
        by_division = (
            df.groupby("事業部")["売上金額"]
            .sum()
            .sort_values(ascending=False)
            .apply(lambda x: f"{x:,}円")
            .to_string()
        )
    by_section = ""
    if "課" in df.columns:
        by_section = (
            df.groupby("課")["売上金額"]
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
    if by_division:
        raw_summary += f"\n【事業部別売上】\n{by_division}\n"
    if by_section:
        raw_summary += f"\n【課別売上】\n{by_section}\n"

    # ── グラフ・表用集計データ ────────────────────────────────────
    df["月"]    = df["日付"].dt.to_period("M").astype(str)
    df["四半期"] = df["日付"].dt.to_period("Q").astype(str)

    # 直近 12 ヶ月の月次合計
    monthly_s     = df.groupby("月")["売上金額"].sum().sort_index()
    monthly_totals = monthly_s.tail(12).to_dict()

    # 商品別合計（降順）
    product_totals = (
        df.groupby("商品名")["売上金額"].sum()
        .sort_values(ascending=False)
        .to_dict()
    )

    # 四半期 × 商品 クロス集計（直近 8 四半期）
    qp = df.pivot_table(
        index="商品名", columns="四半期",
        values="売上金額", aggfunc="sum", fill_value=0,
    )
    if len(qp.columns) > 8:
        qp = qp.iloc[:, -8:]
    quarterly_product_pivot = qp

    # 月次利益率（利益額列がある場合のみ）
    monthly_margin: dict | None = None
    if "利益額" in df.columns:
        m_profit = df.groupby("月")["利益額"].sum()
        m_sales  = df.groupby("月")["売上金額"].sum()
        rate = (m_profit / m_sales * 100).round(1)
        monthly_margin = rate.reindex(list(monthly_totals.keys())).to_dict()

    # 四半期 × 地域 クロス集計（直近 8 四半期）
    qr = df.pivot_table(
        index="地域", columns="四半期",
        values="売上金額", aggfunc="sum", fill_value=0,
    )
    if len(qr.columns) > 8:
        qr = qr.iloc[:, -8:]
    quarterly_region_pivot = qr

    # 四半期 × 担当者 クロス集計（売上合計 上位 8 名、直近 8 四半期）
    rep_totals = df.groupby("担当者")["売上金額"].sum().sort_values(ascending=False)
    top_reps   = rep_totals.index[:8].tolist()
    qrep = df[df["担当者"].isin(top_reps)].pivot_table(
        index="担当者", columns="四半期",
        values="売上金額", aggfunc="sum", fill_value=0,
    )
    # top_reps の順序を維持
    qrep = qrep.reindex([r for r in top_reps if r in qrep.index])
    if len(qrep.columns) > 8:
        qrep = qrep.iloc[:, -8:]
    quarterly_rep_pivot = qrep

    # 年別月次データ（直近3年・月番号キー）
    df["年"]   = df["日付"].dt.year
    df["月番号"] = df["日付"].dt.month
    recent_years = sorted(df["年"].unique())[-3:]
    monthly_by_year: dict = {}
    for y in recent_years:
        ydf = df[df["年"] == y]
        monthly_by_year[int(y)] = {
            int(m): int(v)
            for m, v in ydf.groupby("月番号")["売上金額"].sum().items()
        }

    # 年次前年同期比（YoY）
    yoy_text = ""
    df["年"] = df["日付"].dt.year
    yearly_s = df.groupby("年")["売上金額"].sum().sort_index()
    if len(yearly_s) >= 2:
        yoy_lines = []
        years = list(yearly_s.index)
        for i in range(1, len(years)):
            prev, curr = years[i - 1], years[i]
            prev_val, curr_val = yearly_s[prev], yearly_s[curr]
            if prev_val > 0:
                pct = (curr_val - prev_val) / prev_val * 100
                sign = "+" if pct >= 0 else ""
                yoy_lines.append(f"{curr}年: {sign}{pct:.1f}% (対{prev}年)")
        if yoy_lines:
            yoy_text = "【前年同期比】\n" + "\n".join(yoy_lines) + "\n"
            raw_summary = raw_summary + "\n" + yoy_text

    # 予実比較（売上予定列がある場合のみ）
    budget_by_product: dict | None = None
    actual_vs_budget_text = ""
    if "売上予定" in df.columns:
        actual_s  = df.groupby("商品名")["売上金額"].sum()
        budget_s  = df.groupby("商品名")["売上予定"].sum()
        avb_lines = []
        for prod in actual_s.index:
            act = actual_s[prod]
            bgt = budget_s.get(prod, 0)
            if bgt > 0:
                rate = act / bgt * 100
                avb_lines.append(f"  {prod}: 実績 {act:,.0f}円 / 予定 {bgt:,.0f}円 ({rate:.1f}%)")
        if avb_lines:
            actual_vs_budget_text = "【予実比較（商品別）】\n" + "\n".join(avb_lines) + "\n"
            raw_summary = raw_summary + "\n" + actual_vs_budget_text
        budget_by_product = budget_s.sort_values(ascending=False).to_dict()

    logger.info("集計完了")
    return {
        "total_amount":             total_amount,
        "total_qty":                total_qty,
        "by_product":               by_product,
        "by_region":                by_region,
        "by_rep":                   by_rep,
        "period":                   f"{period_start} ～ {period_end}",
        "raw_summary":              raw_summary,
        "monthly_totals":           monthly_totals,
        "product_totals":           product_totals,
        "quarterly_product_pivot":  quarterly_product_pivot,
        "quarterly_region_pivot":   quarterly_region_pivot,
        "quarterly_rep_pivot":      quarterly_rep_pivot,
        "monthly_margin":           monthly_margin,
        "budget_by_product":        budget_by_product,
        "monthly_by_year":          monthly_by_year,
    }
