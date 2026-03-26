"""
setup_mock.py
ダミーの sales_data.xlsx と template.pptx を data/ ディレクトリに生成するスクリプト。
"""

import os
import random
from datetime import date, timedelta

import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


# ── 出力先 ──────────────────────────────────────────────
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(OUTPUT_DIR, exist_ok=True)

EXCEL_PATH = os.path.join(OUTPUT_DIR, "sales_data.xlsx")
CSV_PATH   = os.path.join(OUTPUT_DIR, "sales_data.csv")
PPTX_PATH  = os.path.join(OUTPUT_DIR, "template.pptx")


# ── 1. sales_data.xlsx ──────────────────────────────────
def create_excel():
    random.seed(42)

    products   = ["商品A", "商品B", "商品C"]
    reps       = ["田中", "佐藤", "鈴木", "山田"]
    regions    = ["東京", "大阪", "名古屋", "福岡"]

    start = date(2025, 1, 1)
    rows  = []
    for i in range(90):
        d      = start + timedelta(days=random.randint(0, 89))
        prod   = random.choice(products)
        rep    = random.choice(reps)
        region = random.choice(regions)
        qty    = random.randint(1, 10)
        unit   = {"商品A": 50000, "商品B": 80000, "商品C": 30000}[prod]
        amount = qty * unit
        rows.append({
            "日付":     d.strftime("%Y-%m-%d"),
            "商品名":   prod,
            "担当者":   rep,
            "地域":     region,
            "数量":     qty,
            "売上金額": amount,
        })

    df = pd.DataFrame(rows).sort_values("日付").reset_index(drop=True)
    df.to_excel(EXCEL_PATH, index=False)
    print(f"[OK] Excel 生成: {EXCEL_PATH}  ({len(df)} 行)")
    df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    print(f"[OK] CSV   生成: {CSV_PATH}  ({len(df)} 行)")


# ── 2. template.pptx ────────────────────────────────────
def _add_textbox(slide, text, left, top, width, height,
                 font_size=18, bold=False, color=None, align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    p  = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    return txBox


def create_pptx():
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]  # blank

    # ── スライド 1：表紙 ──────────────────────────────
    s1 = prs.slides.add_slide(blank_layout)

    # 背景色（濃紺）
    from pptx.oxml.ns import qn
    from lxml import etree
    bg = s1.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0x1F, 0x37, 0x64)

    _add_textbox(s1, "{{report_title}}",
                 left=1, top=2.5, width=11, height=1.2,
                 font_size=36, bold=True,
                 color=(0xFF, 0xFF, 0xFF),
                 align=PP_ALIGN.CENTER)

    _add_textbox(s1, "{{report_date}}",
                 left=1, top=4.0, width=11, height=0.6,
                 font_size=20,
                 color=(0xCC, 0xCC, 0xFF),
                 align=PP_ALIGN.CENTER)

    # ── スライド 2：売上サマリー ───────────────────────
    s2 = prs.slides.add_slide(blank_layout)

    _add_textbox(s2, "売上サマリー",
                 left=0.5, top=0.3, width=12, height=0.7,
                 font_size=28, bold=True,
                 color=(0x1F, 0x37, 0x64))

    # 区切り線（細長いテキストボックスで代用）
    line_box = s2.shapes.add_textbox(
        Inches(0.5), Inches(1.1), Inches(12), Inches(0.05)
    )
    line_box.fill.solid()
    line_box.fill.fore_color.rgb = RGBColor(0x1F, 0x37, 0x64)

    _add_textbox(s2, "{{summary_text}}",
                 left=0.5, top=1.3, width=12, height=5.5,
                 font_size=16)

    # ── スライド 3：所見・次月の方針 ──────────────────
    s3 = prs.slides.add_slide(blank_layout)

    _add_textbox(s3, "所見・次月の方針",
                 left=0.5, top=0.3, width=12, height=0.7,
                 font_size=28, bold=True,
                 color=(0x1F, 0x37, 0x64))

    line_box3 = s3.shapes.add_textbox(
        Inches(0.5), Inches(1.1), Inches(12), Inches(0.05)
    )
    line_box3.fill.solid()
    line_box3.fill.fore_color.rgb = RGBColor(0x1F, 0x37, 0x64)

    _add_textbox(s3, "{{analysis_text}}",
                 left=0.5, top=1.3, width=12, height=5.5,
                 font_size=16)

    prs.save(PPTX_PATH)
    print(f"[OK] PPTX 生成: {PPTX_PATH}  (3 スライド)")


# ── main ─────────────────────────────────────────────────
if __name__ == "__main__":
    print("=== setup_mock.py: ダミーデータ生成 ===")
    create_excel()
    create_pptx()
    print("=== 完了 ===")
