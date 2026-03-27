# -*- coding: utf-8 -*-
"""
텍스트 높이 추정 및 오버플로 감지.

python-pptx는 텍스트 렌더링 측정을 지원하지 않으므로
문자 폭 기반 근사치를 사용합니다.

한글: 전각 (font_size 폭)
영문/숫자: 반각 (font_size * 0.55 폭)
"""

from pptx.util import Inches, Pt, Emu

# 기본 콘텐츠 영역 (와이드스크린 13.333" x 7.5")
DEFAULT_CONTENT_WIDTH = Inches(11.933)   # 13.333 - 0.7*2 margins
DEFAULT_CONTENT_HEIGHT = Inches(5.1)     # 7.5 - 1.4 top - 1.0 title area
DEFAULT_FULL_HEIGHT = Inches(6.3)        # 7.5 - 0.7 top - 0.5 bottom (blank layout)


def estimate_text_height(text: str, font_size_pt: float,
                         content_width_emu: int = None,
                         line_spacing: float = 1.4) -> int:
    """텍스트 높이를 EMU 단위로 추정.

    Args:
        text: 측정할 텍스트
        font_size_pt: 폰트 크기 (pt)
        content_width_emu: 콘텐츠 영역 폭 (EMU). None이면 기본값
        line_spacing: 줄 간격 배수

    Returns:
        추정 높이 (EMU)
    """
    if not text:
        return 0

    if content_width_emu is None:
        content_width_emu = DEFAULT_CONTENT_WIDTH

    font_size_emu = Pt(font_size_pt)

    # 한 줄에 들어가는 문자 수 추정
    chars_per_line = _estimate_chars_per_line(content_width_emu, font_size_emu)
    if chars_per_line <= 0:
        chars_per_line = 1

    # 텍스트 폭 (한글=1.0, 영문=0.55 비율로 계산)
    text_width_units = _text_width_units(text)
    num_lines = max(1, -(-text_width_units // chars_per_line))  # ceil division

    line_height_emu = int(font_size_emu * line_spacing)
    return int(num_lines * line_height_emu)


def estimate_table_height(num_rows: int, num_headers: int = 1,
                          row_height_pt: float = 24) -> int:
    """테이블 높이를 EMU 단위로 추정."""
    total_rows = num_rows + num_headers
    return int(total_rows * Pt(row_height_pt))


def estimate_block_height(block: dict, styles: dict,
                          content_width_emu: int = None) -> int:
    """블록 타입에 따라 높이 추정."""
    block_type = block.get("type", "paragraph")

    if block_type == "bullet":
        level = block.get("level", 1)
        style_key = f"bullet_{level}"
        style = styles.get("styles", {}).get(style_key, styles.get("styles", {}).get("bullet_1", {}))
        text = _runs_to_text(block.get("runs", []))
        font_size = style.get("size", 14)
        spacing = Pt(style.get("space_before_pt", 4) + style.get("space_after_pt", 4))
        return estimate_text_height(text, font_size, content_width_emu,
                                    style.get("line_spacing", 1.4)) + spacing

    elif block_type == "heading":
        return int(Pt(24) * 1.5)  # heading은 한 줄 + 여백

    elif block_type == "paragraph":
        text = _runs_to_text(block.get("runs", []))
        style = styles.get("styles", {}).get("body", {})
        font_size = style.get("size", 12)
        spacing = Pt(style.get("space_before_pt", 4) + style.get("space_after_pt", 4))
        return estimate_text_height(text, font_size, content_width_emu,
                                    style.get("line_spacing", 1.4)) + spacing

    elif block_type == "table":
        rows = block.get("rows", [])
        return estimate_table_height(len(rows))

    elif block_type == "image":
        return int(Inches(4))  # 이미지는 기본 4인치 높이 추정

    return int(Pt(20))  # fallback


def can_fit_blocks(blocks: list, styles: dict,
                   available_height_emu: int,
                   content_width_emu: int = None) -> int:
    """주어진 높이 내에 몇 개의 블록이 들어가는지 반환."""
    total = 0
    for i, block in enumerate(blocks):
        h = estimate_block_height(block, styles, content_width_emu)
        if total + h > available_height_emu and i > 0:
            return i
        total += h
    return len(blocks)


# ---------------------------------------------------------------------------
# 내부 헬퍼
# ---------------------------------------------------------------------------

def _text_width_units(text: str) -> float:
    """텍스트 폭을 단위로 계산. 한글=1.0, 영문=0.55."""
    width = 0.0
    for ch in text:
        if '\u3000' <= ch <= '\u9fff' or '\uac00' <= ch <= '\ud7af':
            width += 1.0  # CJK
        else:
            width += 0.55  # Latin, numbers, symbols
    return width


def _estimate_chars_per_line(content_width_emu: int, font_size_emu: int) -> int:
    """한 줄에 들어가는 문자 수 추정 (한글 기준)."""
    if font_size_emu <= 0:
        return 40
    return int(content_width_emu / font_size_emu)


def _runs_to_text(runs: list) -> str:
    """run 리스트에서 텍스트 추출."""
    return "".join(r.get("text", "") for r in runs)
