# -*- coding: utf-8 -*-
"""
기본 PPTX 템플릿 생성 스크립트.

와이드스크린(13.333" x 7.5") 빈 템플릿을 생성합니다.
기본 레이아웃(Title Slide, Title and Content, Section Header 등)을 그대로 활용합니다.

실행: python create_template.py
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pathlib import Path

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)


def create_default_template():
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    # 기본 레이아웃 이름 재설정 (python-pptx 기본 프레젠테이션에는 11개 레이아웃)
    layout_names = [
        "Title Slide",        # 0
        "Title and Content",  # 1
        "Section Header",     # 2
        "Two Content",        # 3
        "Comparison",         # 4
        "Title Only",         # 5
        "Blank",              # 6
    ]

    layouts = prs.slide_layouts
    for i, name in enumerate(layout_names):
        if i < len(layouts):
            layouts[i].name = name

    # 슬라이드 없이 레이아웃만 가진 빈 템플릿으로 저장
    out_path = Path(__file__).parent / "templates" / "default.pptx"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    print(f"템플릿 생성 완료: {out_path}")
    print(f"레이아웃 수: {len(layouts)}")
    for i, layout in enumerate(layouts):
        phs = ", ".join(f"idx={ph.placeholder_format.idx}({ph.name})" for ph in layout.placeholders)
        print(f"  [{i}] {layout.name}: {phs}")


if __name__ == "__main__":
    create_default_template()
