"""
test_pptx_generator.py — pptx_generator サービスのユニットテスト
"""
import os

import pytest
from pptx import Presentation

from services.pptx_generator import generate_pptx

# プロジェクトルートからテンプレートパスを解決
TEMPLATE_PATH = os.path.join(
    os.path.dirname(__file__), "../../data/template.pptx"
)


@pytest.fixture()
def output_path(tmp_path):
    return str(tmp_path / "output.pptx")


def test_generates_file(output_path):
    generate_pptx(
        template_path=TEMPLATE_PATH,
        output_path=output_path,
        summary_text="テストサマリー",
        analysis_text="テスト分析",
        period="2025/01/01 ～ 2025/03/31",
    )
    assert os.path.exists(output_path)


def test_placeholders_replaced(output_path):
    generate_pptx(
        template_path=TEMPLATE_PATH,
        output_path=output_path,
        summary_text="サマリー本文テキスト",
        analysis_text="分析本文テキスト",
        period="2025/01/01 ～ 2025/03/31",
    )
    prs = Presentation(output_path)
    all_text = " ".join(
        run.text
        for slide in prs.slides
        for shape in slide.shapes
        if shape.has_text_frame
        for para in shape.text_frame.paragraphs
        for run in para.runs
    )
    assert "{{summary_text}}"  not in all_text
    assert "{{analysis_text}}" not in all_text
    assert "{{report_title}}"  not in all_text
    assert "サマリー本文テキスト"  in all_text
    assert "分析本文テキスト"    in all_text


def test_template_not_found(output_path):
    with pytest.raises(FileNotFoundError):
        generate_pptx(
            template_path="/nonexistent/template.pptx",
            output_path=output_path,
            summary_text="",
            analysis_text="",
            period="",
        )
