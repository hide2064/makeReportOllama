"""
create_sogoshosha_sample.py
架空の総合商社「東亜物産株式会社」の売上サンプルデータ生成スクリプト。

実在の総合商社（三菱・三井・伊藤忠など）の事業領域・規模感をベースに、
架空の社名・担当者名を使用して複雑な売上データを生成する。
"""

import csv
import math
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(42)

# ── 商品定義 ────────────────────────────────────────────────────
PRODUCTS = {
    "エネルギー": {
        "items": [
            "原油（中東産）",
            "原油（北海産）",
            "LNG（液化天然ガス）",
            "石炭（一般炭）",
            "石炭（原料炭）",
            "石油製品（ナフサ）",
            "石油製品（重油）",
            "再生可能エネルギー設備",
            "洋上風力発電設備",
        ],
        "base_amount":  (400_000_000, 8_000_000_000),
        "qty_range":    (1, 30),
        "margin":       (0.015, 0.045),
        # 季節係数（冬高・夏低）
        "seasonal": {1: 1.35, 2: 1.25, 3: 1.05, 4: 0.90, 5: 0.80, 6: 0.75,
                     7: 0.85, 8: 0.90, 9: 1.00, 10: 1.10, 11: 1.20, 12: 1.40},
        "trend": 0.005,  # 月次成長率
        "weight": 25,    # 発生頻度ウェイト
    },
    "金属・鉱物": {
        "items": [
            "鉄鋼製品（熱延コイル）",
            "鉄鋼製品（冷延鋼板）",
            "鉄鋼製品（厚板）",
            "アルミニウム地金",
            "アルミニウム板・押出品",
            "銅地金",
            "銅線・銅管",
            "レアアース（ネオジム）",
            "レアアース（ジスプロシウム）",
            "ニッケル地金",
            "亜鉛地金",
            "金属スクラップ",
        ],
        "base_amount":  (50_000_000, 2_000_000_000),
        "qty_range":    (100, 8000),
        "margin":       (0.03, 0.09),
        "seasonal": {m: 1.0 for m in range(1, 13)},
        "trend": 0.003,
        "weight": 20,
    },
    "食料・農産物": {
        "items": [
            "輸入大豆（北米産）",
            "輸入大豆（南米産）",
            "小麦（米国産）",
            "小麦（豪州産）",
            "トウモロコシ（飼料用）",
            "食料加工品（冷凍食品）",
            "食料加工品（缶詰）",
            "水産物（サーモン）",
            "水産物（エビ）",
            "植物油（パーム油）",
            "砂糖（精製糖）",
            "コーヒー豆",
            "カカオ豆",
            "牛肉（冷凍）",
            "豚肉（冷凍）",
        ],
        "base_amount":  (30_000_000, 600_000_000),
        "qty_range":    (10, 2000),
        "margin":       (0.05, 0.14),
        "seasonal": {1: 1.1, 2: 1.0, 3: 1.05, 4: 1.1, 5: 1.0, 6: 0.95,
                     7: 1.0, 8: 1.05, 9: 1.1, 10: 1.15, 11: 1.2, 12: 1.3},
        "trend": 0.008,
        "weight": 18,
    },
    "機械・プラント": {
        "items": [
            "産業機械（工作機械）",
            "産業機械（プレス機）",
            "建設機械（油圧ショベル）",
            "建設機械（クレーン）",
            "発電設備（ガスタービン）",
            "発電設備（変電設備）",
            "半導体製造装置",
            "航空機エンジン部品",
            "船舶用エンジン",
            "鉄道車両部品",
            "医療機器（診断装置）",
            "農業機械",
        ],
        "base_amount":  (80_000_000, 3_000_000_000),
        "qty_range":    (1, 50),
        "margin":       (0.10, 0.25),
        "seasonal": {m: 1.0 for m in range(1, 13)},
        "trend": 0.010,
        "weight": 15,
    },
    "化学品": {
        "items": [
            "エチレン",
            "プロピレン",
            "合成樹脂（PE）",
            "合成樹脂（PP）",
            "合成樹脂（PET）",
            "農薬（除草剤）",
            "農薬（殺虫剤）",
            "化学肥料（尿素）",
            "化学肥料（リン酸）",
            "塗料・コーティング剤",
            "電子材料（半導体用薬品）",
            "医薬品原料",
        ],
        "base_amount":  (20_000_000, 500_000_000),
        "qty_range":    (10, 3000),
        "margin":       (0.05, 0.15),
        "seasonal": {1: 1.0, 2: 1.0, 3: 1.1, 4: 1.1, 5: 1.05, 6: 0.95,
                     7: 0.95, 8: 0.90, 9: 1.0, 10: 1.05, 11: 1.0, 12: 0.95},
        "trend": 0.006,
        "weight": 12,
    },
    "繊維・生活": {
        "items": [
            "繊維原料（綿花）",
            "繊維製品（合成繊維）",
            "アパレル（既製服）",
            "日用消費財",
            "建材（タイル・衛生陶器）",
            "家具・インテリア",
            "スポーツ用品",
            "化粧品・トイレタリー",
        ],
        "base_amount":  (10_000_000, 300_000_000),
        "qty_range":    (100, 10000),
        "margin":       (0.08, 0.20),
        "seasonal": {1: 0.9, 2: 0.85, 3: 1.1, 4: 1.1, 5: 1.0, 6: 0.95,
                     7: 1.0, 8: 1.05, 9: 1.1, 10: 1.1, 11: 1.05, 12: 1.2},
        "trend": 0.004,
        "weight": 10,
    },
}

# ── 地域 ────────────────────────────────────────────────────────
REGIONS = [
    "関東",
    "関西",
    "中部",
    "九州・沖縄",
    "東北",
    "北海道",
    "中国・四国",
    "アジア（東南アジア）",
    "アジア（中国）",
    "アジア（インド）",
    "北米",
    "欧州",
    "中東",
    "オセアニア",
    "中南米",
    "アフリカ",
]

REGION_WEIGHT = [
    15, 13, 10, 6, 4, 3, 4,
    10, 9, 5,
    7, 5, 5, 2, 2, 2,
]

# ── 担当者（架空の名前） ─────────────────────────────────────────
SALES_REPS = [
    # エネルギー担当
    ("田中 誠一",     "エネルギー"),
    ("中村 翔太",     "エネルギー"),
    ("橋本 洋平",     "エネルギー"),
    # 金属担当
    ("佐藤 康弘",     "金属・鉱物"),
    ("加藤 祐樹",     "金属・鉱物"),
    ("岡田 雄介",     "金属・鉱物"),
    # 食料担当
    ("鈴木 大輔",     "食料・農産物"),
    ("小林 裕子",     "食料・農産物"),
    ("高橋 恵",       "食料・農産物"),
    # 機械担当
    ("山田 健一",     "機械・プラント"),
    ("山口 哲也",     "機械・プラント"),
    ("藤原 拓也",     "機械・プラント"),
    # 化学担当
    ("伊藤 浩二",     "化学品"),
    ("吉田 真理",     "化学品"),
    # 繊維・生活担当
    ("渡辺 敏夫",     "繊維・生活"),
    ("石川 美穂",     "繊維・生活"),
    # 海外担当（マルチ商材）
    ("松本 勇気",     None),  # アジア担当
    ("井上 明子",     None),  # 欧州担当
    ("木村 隆史",     None),  # 北米担当
    ("林 俊介",       None),  # 中東担当
    ("青木 健太",     None),  # 中南米担当
]

# ── 日付範囲：2023年1月～2025年3月（2年強） ───────────────────
START_DATE = date(2023, 1, 1)
END_DATE   = date(2025, 3, 31)

def random_date_in_month(year: int, month: int) -> date:
    import calendar
    last_day = calendar.monthrange(year, month)[1]
    return date(year, month, random.randint(1, last_day))

def generate_rows() -> list[dict]:
    rows = []

    # 月ごとに生成
    current = date(START_DATE.year, START_DATE.month, 1)
    month_index = 0

    while current <= END_DATE:
        y, m = current.year, current.month
        # 総取引件数（月により変動、期末は多め）
        is_quarter_end = m in (3, 6, 9, 12)
        base_txn = 25 if is_quarter_end else 18
        n_transactions = random.randint(base_txn - 4, base_txn + 6)

        for _ in range(n_transactions):
            # カテゴリをウェイト付きでランダム選択
            categories  = list(PRODUCTS.keys())
            cat_weights = [PRODUCTS[c]["weight"] for c in categories]
            category    = random.choices(categories, weights=cat_weights, k=1)[0]
            pdef        = PRODUCTS[category]

            product = random.choice(pdef["items"])

            # 担当者選択（専門担当 or 海外担当）
            specialists = [r for r in SALES_REPS if r[1] == category]
            generalists = [r for r in SALES_REPS if r[1] is None]
            pool = specialists * 3 + generalists  # 専門担当を優先
            rep_name, _ = random.choice(pool)

            # 地域選択
            region = random.choices(REGIONS, weights=REGION_WEIGHT, k=1)[0]

            # 売上金額（季節係数＋トレンド）
            seasonal_factor = pdef["seasonal"].get(m, 1.0)
            trend_factor    = (1 + pdef["trend"]) ** month_index
            base_lo, base_hi = pdef["base_amount"]
            base_amount = random.uniform(base_lo, base_hi)
            amount = int(base_amount * seasonal_factor * trend_factor)
            # 1,000円単位に丸める
            amount = round(amount / 1000) * 1000

            # 数量
            qty = random.randint(*pdef["qty_range"])

            # 利益額
            margin_rate = random.uniform(*pdef["margin"])
            # 海外取引は利益率やや高め
            if region in ("北米", "欧州", "中東", "中南米", "アフリカ", "アジア（東南アジア）"):
                margin_rate *= random.uniform(1.05, 1.20)
            profit = int(amount * margin_rate)

            txn_date = random_date_in_month(y, m)

            rows.append({
                "日付":     txn_date.strftime("%Y/%m/%d"),
                "商品名":   product,
                "カテゴリ": category,
                "担当者":   rep_name,
                "地域":     region,
                "数量":     qty,
                "売上金額": amount,
                "利益額":   profit,
            })

        # 月を進める
        if m == 12:
            current = date(y + 1, 1, 1)
        else:
            current = date(y, m + 1, 1)
        month_index += 1

    # 日付順にソート
    rows.sort(key=lambda r: r["日付"])
    return rows


def main():
    rows = generate_rows()
    out_path = Path(__file__).parent / "data" / "sample_sogoshosha_eastasia.csv"
    out_path.parent.mkdir(exist_ok=True)

    fieldnames = ["日付", "商品名", "カテゴリ", "担当者", "地域", "数量", "売上金額", "利益額"]
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    # 統計サマリーを表示
    total = sum(r["売上金額"] for r in rows)
    total_profit = sum(r["利益額"] for r in rows)
    print(f"生成完了: {out_path}")
    print(f"  レコード数  : {len(rows):,} 件")
    print(f"  期間        : {rows[0]['日付']} ～ {rows[-1]['日付']}")
    print(f"  総売上金額  : {total:,} 円")
    print(f"  総利益額    : {total_profit:,} 円")
    print(f"  平均利益率  : {total_profit / total * 100:.1f}%")

    # カテゴリ別サマリー
    from collections import defaultdict
    by_cat = defaultdict(int)
    for r in rows:
        by_cat[r["カテゴリ"]] += r["売上金額"]
    print("\n  カテゴリ別売上:")
    for cat, amt in sorted(by_cat.items(), key=lambda x: -x[1]):
        print(f"    {cat:<16}: {amt:>20,} 円  ({amt/total*100:.1f}%)")

    # 地域別サマリー
    by_region = defaultdict(int)
    for r in rows:
        by_region[r["地域"]] += r["売上金額"]
    print("\n  地域別売上 TOP 5:")
    for reg, amt in sorted(by_region.items(), key=lambda x: -x[1])[:5]:
        print(f"    {reg:<20}: {amt:>20,} 円")


if __name__ == "__main__":
    main()
