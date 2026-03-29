"""
create_complex_samples.py
2019〜2025年の7年間にわたる複雑な売上サンプルデータを生成する。

生成ファイル:
  data/sample_complex.xlsx   — フル Excel（全列）
  data/sample_complex.csv    — CSV 版
  data/sample_complex_2024.xlsx — 直近1年抜粋（動作確認用）
"""

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(2024)

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

# ──────────────────────────────────────────────
# マスターデータ定義
# ──────────────────────────────────────────────

# 商品マスター: (商品名, カテゴリ, 単価帯, 原価率, 需要傾向)
# 需要傾向: 1〜12月の相対値
PRODUCTS = [
    # ── SaaS プラン ──
    ("エンタープライズプラン",  "SaaS",      320_000, 0.28,
     [0.8,0.8,1.0,1.1,1.1,1.0,0.9,1.0,1.2,1.2,1.3,1.6]),
    ("プレミアムプラン",        "SaaS",      150_000, 0.35,
     [0.9,0.9,1.0,1.0,1.1,1.0,1.0,1.1,1.2,1.2,1.3,1.5]),
    ("スタンダードプラン",      "SaaS",       60_000, 0.48,
     [1.0,0.9,1.0,1.0,1.0,1.1,1.1,1.0,1.0,1.1,1.2,1.4]),
    ("ライトプラン",            "SaaS",       18_000, 0.55,
     [1.0,1.0,1.1,1.0,1.0,1.0,1.0,0.9,1.0,1.0,1.1,1.2]),
    # ── コンサル ──
    ("戦略コンサルティング",    "Consulting", 500_000, 0.25,
     [0.6,0.7,1.1,1.2,1.1,1.0,0.8,0.9,1.2,1.3,1.1,1.2]),
    ("業務改善コンサルティング","Consulting", 250_000, 0.30,
     [0.7,0.8,1.0,1.1,1.1,1.0,0.9,1.0,1.2,1.2,1.1,1.3]),
    ("IT導入支援",              "Consulting", 180_000, 0.38,
     [0.8,0.9,1.0,1.0,1.1,1.0,1.0,1.0,1.1,1.1,1.2,1.3]),
    # ── 保守・サポート ──
    ("プレミアムサポート",      "Support",    45_000, 0.22,
     [1.0,1.0,1.1,1.0,1.0,1.0,1.0,1.0,1.1,1.0,1.0,1.1]),
    ("スタンダードサポート",    "Support",    20_000, 0.28,
     [1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0]),
    ("ヘルプデスク",            "Support",     8_000, 0.35,
     [1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0,1.0]),
    # ── ハードウェア/ライセンス ──
    ("サーバーライセンス",      "License",   400_000, 0.55,
     [1.1,1.0,1.2,1.0,1.1,0.9,0.9,1.0,1.1,1.2,1.1,1.3]),
    ("クライアントライセンス",  "License",    35_000, 0.60,
     [1.0,1.0,1.1,1.0,1.0,0.9,0.9,1.0,1.1,1.1,1.1,1.4]),
]

# 地域マスター: (地域名, 経済圏ウェイト)
REGIONS = [
    ("東京",  0.32),
    ("大阪",  0.18),
    ("名古屋",0.13),
    ("福岡",  0.10),
    ("札幌",  0.07),
    ("仙台",  0.07),
    ("広島",  0.07),
    ("那覇",  0.06),
]

# 担当者マスター: (名前, 拠点, 得意カテゴリ, 能力スコア)
REPS = [
    ("田中 一郎",    "東京",  "SaaS",        1.15),
    ("佐藤 花子",    "大阪",  "Consulting",  1.20),
    ("鈴木 太郎",    "東京",  "License",     1.05),
    ("伊藤 美咲",    "名古屋","SaaS",        1.10),
    ("渡辺 健二",    "福岡",  "Support",     1.00),
    ("中村 さくら",  "東京",  "Consulting",  1.25),
    ("小林 大輔",    "大阪",  "License",     1.08),
    ("加藤 里奈",    "東京",  "SaaS",        1.12),
    ("吉田 拓海",    "名古屋","Consulting",  1.18),
    ("山田 彩香",    "福岡",  "SaaS",        0.98),
    ("松本 浩二",    "札幌",  "Support",     1.02),
    ("井上 真由",    "仙台",  "License",     1.06),
    ("木村 隼人",    "東京",  "Consulting",  1.22),
    ("林 奈緒",      "大阪",  "SaaS",        1.14),
    ("清水 俊介",    "広島",  "Support",     0.97),
    ("山口 恵子",    "東京",  "SaaS",        1.09),
    ("森 賢太",      "名古屋","License",     1.03),
    ("池田 美穂",    "大阪",  "Consulting",  1.16),
    ("橋本 雄大",    "那覇",  "SaaS",        1.01),
    ("石川 沙織",    "東京",  "License",     1.07),
]

# 顧客タイプ
CUSTOMER_TYPES = [
    ("大手企業",     0.20, 1.8),   # (タイプ名, 確率, 単価倍率)
    ("中堅企業",     0.35, 1.2),
    ("中小企業",     0.30, 0.9),
    ("スタートアップ",0.10, 0.7),
    ("官公庁",       0.05, 1.5),
]

# 販売チャネル
CHANNELS = [
    ("直販",    0.45),
    ("代理店",  0.35),
    ("Web",     0.15),
    ("紹介",    0.05),
]

# キャンペーン（月・名前・売上インパクト）
CAMPAIGNS = {
    (3,  "年度末キャンペーン"):   1.35,
    (6,  "上半期フェア"):         1.20,
    (9,  "秋季セール"):           1.25,
    (12, "年末特別割引"):         1.40,
}

# 年次成長率（2019=1.0 基準）
YEARLY_GROWTH = {
    2019: 1.00,
    2020: 0.85,   # コロナ影響
    2021: 0.98,   # 回復途中
    2022: 1.12,
    2023: 1.24,
    2024: 1.35,
    2025: 1.18,   # 直近（データが少ない）
}

# 2025年は1〜3月のみ生成
YEAR_MONTHS = {y: range(1, 13) for y in range(2019, 2025)}
YEAR_MONTHS[2025] = range(1, 4)


def pick(items, weights=None):
    if weights is None:
        return random.choice(items)
    return random.choices(items, weights=weights, k=1)[0]


def gen_rows(year: int, month: int) -> list[dict]:
    rows = []
    growth     = YEARLY_GROWTH[year]
    # 月ごとの基本取引件数（2019年基準で月20〜35件）
    base_count = random.randint(20, 35)
    count      = max(8, int(base_count * growth))

    # キャンペーン効果
    campaign_key  = next((k for k in CAMPAIGNS if k[0] == month), None)
    campaign_name = campaign_key[1] if campaign_key else ""
    campaign_mult = CAMPAIGNS[campaign_key] if campaign_key else 1.0

    for _ in range(count):
        # 商品選択
        prod_weights = [0.15, 0.18, 0.22, 0.10, 0.05, 0.07, 0.06, 0.06, 0.04, 0.03, 0.02, 0.02]
        prod = pick(PRODUCTS, prod_weights)
        pname, category, base_price, cost_rate, seasonality = prod

        # 顧客タイプ選択
        ct_names  = [c[0] for c in CUSTOMER_TYPES]
        ct_probs  = [c[1] for c in CUSTOMER_TYPES]
        ct_mults  = [c[2] for c in CUSTOMER_TYPES]
        ct_idx    = CUSTOMER_TYPES.index(pick(CUSTOMER_TYPES, ct_probs))
        cust_type = ct_names[ct_idx]
        price_mult= ct_mults[ct_idx]

        # チャネル選択
        ch_names = [c[0] for c in CHANNELS]
        ch_probs = [c[1] for c in CHANNELS]
        channel  = pick(CHANNELS, ch_probs)[0]

        # 担当者（カテゴリ適性を考慮）
        rep_weights = [
            2.5 if r[2] == category else 1.0
            for r in REPS
        ]
        rep = pick(REPS, rep_weights)
        rep_name, rep_base, _, rep_score = rep

        # 地域（担当者拠点 or ランダム）
        region_weights = []
        for rname, rw in REGIONS:
            w = rw * (3.0 if rname == rep_base else 1.0)
            region_weights.append(w)
        region = pick(REGIONS, region_weights)[0]

        # 数量
        if base_price >= 200_000:
            qty = random.randint(1, 3)
        elif base_price >= 50_000:
            qty = random.randint(1, 8)
        else:
            qty = random.randint(1, 30)

        # 売上金額計算
        seasonal = seasonality[month - 1]
        noise    = random.uniform(0.85, 1.15)
        amount   = int(base_price * qty * seasonal * growth * price_mult
                       * rep_score * campaign_mult * noise)
        amount   = max(5_000, amount)  # 最小保証

        # 原価・利益
        cost_noise = random.uniform(0.93, 1.07)
        cost   = int(amount * cost_rate * cost_noise)
        profit = amount - cost
        margin = round(profit / amount * 100, 1) if amount > 0 else 0.0

        # 日付（月内のランダム営業日）
        max_day = {2: 28}.get(month, 30) if year % 4 != 0 else {2: 29}.get(month, 30)
        max_day = min(max_day, 28)
        day = random.randint(1, max_day)
        dt  = date(year, month, day).strftime("%Y-%m-%d")

        # 顧客ID（大手は固定顧客、SMB はランダム）
        if cust_type == "大手企業":
            cust_id = f"ENT-{random.randint(1, 30):04d}"
        elif cust_type == "官公庁":
            cust_id = f"GOV-{random.randint(1, 15):04d}"
        else:
            cust_id = f"SMB-{random.randint(1, 300):05d}"

        rows.append({
            "日付":       dt,
            "商品名":     pname,
            "カテゴリ":   category,
            "担当者":     rep_name,
            "地域":       region,
            "顧客タイプ": cust_type,
            "顧客ID":     cust_id,
            "チャネル":   channel,
            "数量":       qty,
            "売上金額":   amount,
            "原価":       cost,
            "利益額":     profit,
            "利益率(%)":  margin,
            "キャンペーン": campaign_name,
        })

    return rows


def main():
    all_rows = []
    for year, months in YEAR_MONTHS.items():
        for month in months:
            rows = gen_rows(year, month)
            all_rows.extend(rows)
            print(f"  {year}/{month:02d}: {len(rows):3d} 件 生成")

    df = pd.DataFrame(all_rows).sort_values("日付").reset_index(drop=True)

    # ── 全件ファイル ──
    out_xlsx = DATA_DIR / "sample_complex.xlsx"
    out_csv  = DATA_DIR / "sample_complex.csv"
    df.to_excel(out_xlsx, index=False)
    df.to_csv(out_csv,    index=False, encoding="utf-8-sig")

    # ── 2024年抜粋（直近1年・動作確認用） ──
    df_2024 = df[df["日付"].str.startswith("2024")].copy()
    out_2024 = DATA_DIR / "sample_complex_2024.xlsx"
    df_2024.to_excel(out_2024, index=False)

    # ── 統計サマリー表示 ──
    print()
    print("=" * 60)
    print(f"全件: {len(df):,} 行  /  {df['売上金額'].sum():,.0f}円")
    print(f"期間: {df['日付'].min()} 〜 {df['日付'].max()}")
    print()
    print("[年次集計]")
    yearly = df.groupby(df["日付"].str[:4].rename("年"))["売上金額"].sum()
    for yr, amt in yearly.items():
        print(f"  {yr}年: {amt:>15,.0f}円")
    print()
    print("[カテゴリ別]")
    for cat, amt in df.groupby("カテゴリ")["売上金額"].sum().sort_values(ascending=False).items():
        print(f"  {cat:20s}: {amt:>15,.0f}円")
    print()
    print("[地域別 Top5]")
    for reg, amt in df.groupby("地域")["売上金額"].sum().sort_values(ascending=False).head(5).items():
        print(f"  {reg:10s}: {amt:>15,.0f}円")
    print()
    print(f"出力ファイル:")
    print(f"  {out_xlsx}   ({len(df):,} 行)")
    print(f"  {out_csv}")
    print(f"  {out_2024}   ({len(df_2024):,} 行 / 2024年のみ)")
    print("=" * 60)


if __name__ == "__main__":
    main()
