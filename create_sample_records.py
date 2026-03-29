"""
create_sample_records.py
5種類のサンプル売上データファイルを生成する。

生成ファイル:
  data/sample_01_startup.xlsx    — スタートアップ (2024年, 小規模, SaaS偏重)
  data/sample_02_regional.xlsx   — 地方中堅企業 (2023-2024年, 小売・卸売)
  data/sample_03_enterprise.xlsx — 大手企業 (2022-2024年, 多カテゴリ, 高単価)
  data/sample_04_manufacturing.xlsx — 製造業 (2024年, ハードウェア+保守)
  data/sample_05_consulting.xlsx — コンサルファーム (2023-2024年, 案件型)
"""

import random
from datetime import date
from pathlib import Path

import pandas as pd

random.seed(42)
DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)


def rnd_date(year: int, month: int) -> str:
    day = random.randint(1, 28)
    return date(year, month, day).strftime("%Y-%m-%d")


def make_row(dt, product, category, rep, region,
             cust_type, cust_id, channel, qty, price,
             cost_rate, campaign=""):
    amount = int(price * qty * random.uniform(0.88, 1.12))
    amount = max(5_000, amount)
    cost   = int(amount * cost_rate * random.uniform(0.93, 1.07))
    profit = amount - cost
    margin = round(profit / amount * 100, 1) if amount > 0 else 0.0
    return {
        "日付": dt,
        "商品名": product,
        "カテゴリ": category,
        "担当者": rep,
        "地域": region,
        "顧客タイプ": cust_type,
        "顧客ID": cust_id,
        "チャネル": channel,
        "数量": qty,
        "売上金額": amount,
        "原価": cost,
        "利益額": profit,
        "利益率(%)": margin,
        "キャンペーン": campaign,
    }


# ══════════════════════════════════════════════════════════════
# 1. スタートアップ (2024年, SaaS偏重, 小規模 ~180件)
# ══════════════════════════════════════════════════════════════
def gen_startup():
    products = [
        ("フリープランUpgrade", "SaaS", 9_800, 0.40),
        ("スタータープラン",    "SaaS", 29_800, 0.38),
        ("グロースプラン",      "SaaS", 79_800, 0.32),
        ("エンタープライズ",    "SaaS", 198_000, 0.25),
        ("オンボーディング支援","Service", 150_000, 0.35),
        ("API利用従量課金",     "SaaS", 15_000, 0.20),
    ]
    reps = ["山田 翔太", "鈴木 あかり", "佐藤 健", "田中 みゆ"]
    regions = ["東京", "大阪", "福岡", "その他"]
    cust_types = [("スタートアップ", 0.45, 0.8), ("中小企業", 0.35, 1.0),
                  ("中堅企業", 0.15, 1.3), ("大手企業", 0.05, 1.8)]
    channels = [("Web", 0.55), ("直販", 0.30), ("紹介", 0.15)]
    campaign_months = {3: "春のスタートダッシュ", 6: "夏割キャンペーン",
                       9: "秋の成長支援", 12: "年末特典"}

    rows = []
    for month in range(1, 13):
        base = random.randint(12, 20)
        camp = campaign_months.get(month, "")
        mult = 1.3 if camp else 1.0
        count = int(base * mult)
        for _ in range(count):
            p = random.choices(products, weights=[0.20, 0.30, 0.25, 0.08, 0.07, 0.10])[0]
            ct = random.choices(cust_types, weights=[c[1] for c in cust_types])[0]
            qty = random.randint(1, 5) if p[2] < 50_000 else 1
            cid = f"SU-{random.randint(1, 200):04d}"
            ch = random.choices([c[0] for c in channels], weights=[c[1] for c in channels])[0]
            row = make_row(
                rnd_date(2024, month), p[0], p[1],
                random.choice(reps), random.choice(regions),
                ct[0], cid, ch, qty, int(p[2] * ct[2]), p[3], camp
            )
            rows.append(row)
    return rows


# ══════════════════════════════════════════════════════════════
# 2. 地方中堅小売・卸売 (2023-2024年, ~320件)
# ══════════════════════════════════════════════════════════════
def gen_regional():
    products = [
        ("食品・飲料",    "食品", 3_500, 0.62),
        ("日用品",        "日用品", 2_800, 0.58),
        ("電化製品",      "家電", 45_000, 0.55),
        ("農産物（直送）","食品", 8_000, 0.45),
        ("地域特産品",    "特産品", 12_000, 0.40),
        ("業務用食材",    "食品", 25_000, 0.52),
        ("包装・資材",    "資材", 18_000, 0.60),
    ]
    reps = ["本田 一郎", "中村 幸子", "西村 拓", "松田 恵", "岡田 太郎"]
    regions = ["広島", "岡山", "山口", "鳥取", "島根"]
    cust_types = [("中小企業", 0.40, 1.0), ("個人事業主", 0.30, 0.85),
                  ("中堅企業", 0.20, 1.2), ("官公庁", 0.10, 1.3)]
    channels = [("直販", 0.50), ("代理店", 0.35), ("Web", 0.15)]
    campaign_months = {2: "春節セール", 8: "お盆特売", 11: "年末商戦準備"}

    rows = []
    for year in [2023, 2024]:
        growth = 1.0 if year == 2023 else 1.08
        for month in range(1, 13):
            base = random.randint(10, 18)
            camp = campaign_months.get(month, "")
            mult = (1.25 if camp else 1.0) * growth
            count = int(base * mult)
            for _ in range(count):
                p = random.choices(products, weights=[0.22, 0.18, 0.12, 0.15, 0.12, 0.13, 0.08])[0]
                ct = random.choices(cust_types, weights=[c[1] for c in cust_types])[0]
                qty = random.randint(5, 100) if p[2] < 10_000 else random.randint(1, 10)
                cid = f"REG-{random.randint(1, 150):04d}"
                ch = random.choices([c[0] for c in channels], weights=[c[1] for c in channels])[0]
                row = make_row(
                    rnd_date(year, month), p[0], p[1],
                    random.choice(reps), random.choice(regions),
                    ct[0], cid, ch, qty, int(p[2] * ct[2]), p[3], camp
                )
                rows.append(row)
    return rows


# ══════════════════════════════════════════════════════════════
# 3. 大手エンタープライズ (2022-2024年, 多カテゴリ, ~480件)
# ══════════════════════════════════════════════════════════════
def gen_enterprise():
    products = [
        ("エンタープライズERP",  "Software", 5_000_000, 0.20),
        ("クラウド移行支援",     "Consulting", 3_000_000, 0.25),
        ("セキュリティ監査",     "Consulting", 1_500_000, 0.28),
        ("年間保守契約",         "Support", 800_000, 0.18),
        ("データ分析基盤構築",   "Software", 2_500_000, 0.22),
        ("トレーニング・研修",   "Service", 300_000, 0.35),
        ("ライセンス（追加）",   "License", 500_000, 0.45),
        ("ヘルプデスク運用",     "Support", 200_000, 0.30),
    ]
    reps = ["木村 誠司", "橋本 彩子", "渡辺 浩二", "斎藤 里奈",
            "伊藤 大介", "小林 美咲", "中島 健太"]
    regions = ["東京", "大阪", "名古屋", "福岡", "仙台", "広島"]
    cust_types = [("大手企業", 0.55, 1.8), ("官公庁", 0.25, 1.5),
                  ("中堅企業", 0.15, 1.2), ("外資系", 0.05, 2.0)]
    channels = [("直販", 0.70), ("代理店", 0.20), ("紹介", 0.10)]
    campaign_months = {3: "年度末商談", 9: "下期キックオフ", 12: "年末予算消化"}
    growth_by_year = {2022: 1.0, 2023: 1.12, 2024: 1.22}

    rows = []
    for year in [2022, 2023, 2024]:
        growth = growth_by_year[year]
        for month in range(1, 13):
            base = random.randint(6, 12)
            camp = campaign_months.get(month, "")
            mult = (1.4 if camp else 1.0) * growth
            count = int(base * mult)
            for _ in range(count):
                p = random.choices(products, weights=[0.12, 0.15, 0.12, 0.18, 0.12, 0.10, 0.11, 0.10])[0]
                ct = random.choices(cust_types, weights=[c[1] for c in cust_types])[0]
                qty = 1
                cid = f"ENT-{random.randint(1, 50):04d}"
                ch = random.choices([c[0] for c in channels], weights=[c[1] for c in channels])[0]
                row = make_row(
                    rnd_date(year, month), p[0], p[1],
                    random.choice(reps), random.choice(regions),
                    ct[0], cid, ch, qty, int(p[2] * ct[2]), p[3], camp
                )
                rows.append(row)
    return rows


# ══════════════════════════════════════════════════════════════
# 4. 製造業（ハードウェア＋保守） (2024年, ~220件)
# ══════════════════════════════════════════════════════════════
def gen_manufacturing():
    products = [
        ("産業用センサー（標準）", "ハードウェア", 85_000, 0.55),
        ("産業用センサー（高精度）","ハードウェア", 240_000, 0.50),
        ("制御ユニット",           "ハードウェア", 320_000, 0.52),
        ("IoTゲートウェイ",        "ハードウェア", 150_000, 0.48),
        ("設置・導入工事",         "工事", 200_000, 0.38),
        ("年間保守（標準）",       "保守", 60_000, 0.22),
        ("年間保守（プレミアム）", "保守", 120_000, 0.18),
        ("交換部品・消耗品",       "部品", 15_000, 0.60),
        ("クラウド監視サービス",   "SaaS", 45_000, 0.25),
    ]
    reps = ["高橋 剛", "山本 直美", "藤田 誠", "石田 麻衣", "前田 浩"]
    regions = ["愛知", "静岡", "神奈川", "大阪", "広島", "北九州"]
    cust_types = [("大手企業", 0.30, 1.6), ("中堅企業", 0.45, 1.1),
                  ("中小企業", 0.20, 0.9), ("官公庁", 0.05, 1.4)]
    channels = [("直販", 0.60), ("代理店", 0.35), ("Web", 0.05)]
    campaign_months = {4: "新年度導入促進", 10: "冬季メンテナンス特価"}

    rows = []
    for month in range(1, 13):
        base = random.randint(14, 22)
        camp = campaign_months.get(month, "")
        mult = 1.3 if camp else 1.0
        count = int(base * mult)
        for _ in range(count):
            p = random.choices(products, weights=[0.18, 0.10, 0.08, 0.10, 0.12, 0.15, 0.08, 0.12, 0.07])[0]
            ct = random.choices(cust_types, weights=[c[1] for c in cust_types])[0]
            qty = random.randint(1, 20) if p[2] < 30_000 else random.randint(1, 5)
            cid = f"MFG-{random.randint(1, 120):04d}"
            ch = random.choices([c[0] for c in channels], weights=[c[1] for c in channels])[0]
            row = make_row(
                rnd_date(2024, month), p[0], p[1],
                random.choice(reps), random.choice(regions),
                ct[0], cid, ch, qty, int(p[2] * ct[2]), p[3], camp
            )
            rows.append(row)
    return rows


# ══════════════════════════════════════════════════════════════
# 5. コンサルファーム (2023-2024年, 案件型, ~200件)
# ══════════════════════════════════════════════════════════════
def gen_consulting():
    products = [
        ("DX戦略策定",         "戦略", 3_500_000, 0.22),
        ("業務プロセス改革",   "BPR", 2_000_000, 0.26),
        ("IT戦略ロードマップ", "戦略", 1_800_000, 0.24),
        ("PMO支援",            "PM", 800_000, 0.30),
        ("組織変革コンサル",   "人材", 1_200_000, 0.28),
        ("データ活用支援",     "データ", 900_000, 0.27),
        ("M&Aデューデリ",      "M&A", 2_500_000, 0.20),
        ("研修・ワークショップ","教育", 400_000, 0.38),
        ("調査・レポート作成", "調査", 600_000, 0.35),
    ]
    reps = ["吉村 正一", "竹内 綾子", "河野 裕樹", "森田 千夏",
            "坂本 賢二", "長谷川 亮"]
    regions = ["東京", "大阪", "名古屋", "福岡", "海外（アジア）"]
    cust_types = [("大手企業", 0.50, 1.8), ("中堅企業", 0.30, 1.2),
                  ("官公庁", 0.15, 1.5), ("外資系", 0.05, 2.2)]
    channels = [("直販", 0.65), ("紹介", 0.25), ("代理店", 0.10)]
    campaign_months = {3: "年度末提案ラッシュ", 9: "下期予算獲得支援"}
    growth_by_year = {2023: 1.0, 2024: 1.15}

    rows = []
    for year in [2023, 2024]:
        growth = growth_by_year[year]
        for month in range(1, 13):
            base = random.randint(5, 10)
            camp = campaign_months.get(month, "")
            mult = (1.5 if camp else 1.0) * growth
            count = int(base * mult)
            for _ in range(count):
                p = random.choices(products, weights=[0.12, 0.14, 0.11, 0.13, 0.10, 0.12, 0.08, 0.12, 0.08])[0]
                ct = random.choices(cust_types, weights=[c[1] for c in cust_types])[0]
                qty = 1
                cid = f"CSL-{random.randint(1, 80):04d}"
                ch = random.choices([c[0] for c in channels], weights=[c[1] for c in channels])[0]
                row = make_row(
                    rnd_date(year, month), p[0], p[1],
                    random.choice(reps), random.choice(regions),
                    ct[0], cid, ch, qty, int(p[2] * ct[2]), p[3], camp
                )
                rows.append(row)
    return rows


# ══════════════════════════════════════════════════════════════
# メイン: 5ファイル出力
# ══════════════════════════════════════════════════════════════
SAMPLES = [
    ("sample_01_startup",      "スタートアップ (2024年)",           gen_startup),
    ("sample_02_regional",     "地方中堅小売・卸売 (2023-2024年)",  gen_regional),
    ("sample_03_enterprise",   "大手エンタープライズ (2022-2024年)",gen_enterprise),
    ("sample_04_manufacturing","製造業ハードウェア (2024年)",       gen_manufacturing),
    ("sample_05_consulting",   "コンサルファーム (2023-2024年)",    gen_consulting),
]


# コンサル向けシンプル列構成（8列）
SIMPLE_COLS = ["日付", "商品名", "カテゴリ", "担当者", "地域", "数量", "売上金額", "利益率(%)"]


def main():
    print("=" * 65)
    print("サンプルレコードファイル生成")
    print("=" * 65)

    for stem, label, gen_fn in SAMPLES:
        rows = gen_fn()
        df_full = pd.DataFrame(rows).sort_values("日付").reset_index(drop=True)

        # フル版（全14列）
        out_xlsx = DATA_DIR / f"{stem}.xlsx"
        out_csv  = DATA_DIR / f"{stem}.csv"
        df_full.to_excel(out_xlsx, index=False)
        df_full.to_csv(out_csv, index=False, encoding="utf-8-sig")

        # シンプル版（8列・コンサル現場向け）
        df_simple = df_full[SIMPLE_COLS].copy()
        out_simple = DATA_DIR / f"{stem}_simple.xlsx"
        out_simple_csv = DATA_DIR / f"{stem}_simple.csv"
        df_simple.to_excel(out_simple, index=False)
        df_simple.to_csv(out_simple_csv, index=False, encoding="utf-8-sig")

        total = df_full["売上金額"].sum()
        print(f"\n[{stem}]  {label}")
        print(f"  件数: {len(df_full):,} 行 / 売上合計: {total:,.0f} 円")
        print(f"  期間: {df_full['日付'].min()} 〜 {df_full['日付'].max()}")
        print(f"  フル版:    {out_xlsx.name}")
        print(f"  シンプル版: {out_simple.name}  ({', '.join(SIMPLE_COLS)})")

    print("\n" + "=" * 65)
    print("完了: 5ファイル x 4形式 (full xlsx/csv + simple xlsx/csv) = 20ファイル")
    print("=" * 65)


if __name__ == "__main__":
    main()
