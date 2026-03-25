"""
pptx_generator.py
テンプレート PPTX のプレースホルダーにテキストを埋め込んで出力する。
"""

import logging
from copy import deepcopy
from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.util import Pt

logger = logging.getLogger(__name__)

PLACEHOLDER_MAP = {
    "{{report_title}}": None,   # 呼び出し元で設定
    "{{report_date}}":  None,
    "{{summary_text}}": None,
    "{{analysis_text}}": None,
}


def _replace_text_in_slide(slide, replacements: dict[str, str]):
    """スライド内のすべてのテキストボックスに対してプレースホルダー置換を行う。"""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                for placeholder, value in replacements.items():
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, value)
                        logger.debug(f"置換: {placeholder} → (len={len(value)})")


def generate_pptx(
    template_path: str,
    output_path:   str,
    summary_text:  str,
    analysis_text: str,
    period:        str,
) -> str:
    """
    テンプレートを元に報告書 PPTX を生成して output_path に保存。
    生成したファイルパスを返す。
    """
    logger.info(f"PPTX 生成開始: template={template_path}")
    if not Path(template_path).exists():
        raise FileNotFoundError(f"テンプレートが見つかりません: {template_path}")

    prs = Presentation(template_path)

    today  = date.today().strftime("%Y年%m月%d日")
    title  = f"月次売上報告書（{period}）"

    replacements = {
        "{{report_title}}":  title,
        "{{report_date}}":   f"作成日: {today}",
        "{{summary_text}}":  summary_text,
        "{{analysis_text}}": analysis_text,
    }

    for slide in prs.slides:
        _replace_text_in_slide(slide, replacements)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    logger.info(f"PPTX 保存完了: {output_path}")
    return output_path
