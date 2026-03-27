# -*- coding: utf-8 -*-
"""
SlideTree → SlideSpec[] 변환 (직접변환 모드용)

일반 마크다운의 H1-H5 계층 트리를 슬라이드 단위로 분할합니다.
세션 모드에서는 Claude가 이미 슬라이드를 나눠주므로 이 모듈은 사용하지 않습니다.
"""

from text_fitting import (
    estimate_block_height,
    can_fit_blocks,
    DEFAULT_CONTENT_HEIGHT,
    DEFAULT_CONTENT_WIDTH,
)
from md_parser import parse_inline_runs


def plan_slides(tree: dict, styles: dict) -> list:
    """마크다운 트리를 SlideSpec 리스트로 변환.

    Args:
        tree: parse_markdown_tree()의 반환값
        styles: pptx-styles.json 내용

    Returns:
        SlideSpec 리스트
    """
    slides = []
    overflow = styles.get("overflow", {})
    max_table_rows = overflow.get("max_table_rows_per_slide", 15)

    for part in tree.get("parts", []):
        # H1 → Section Header
        if part["title"]:
            slides.append({
                "layout": "section_header",
                "title": part["title"],
                "subtitle": "",
                "blocks": [],
            })

        for chapter in part.get("chapters", []):
            # H2 → 챕터 시작
            for section in chapter.get("sections", []):
                # 블록들을 슬라이드에 분배
                blocks = _preprocess_blocks(section.get("blocks", []), max_table_rows)

                if not blocks and not section["title"]:
                    continue

                # 블록을 슬라이드 높이에 맞게 분할
                split_slides = _split_blocks_to_slides(
                    blocks,
                    chapter["title"],
                    section["title"],
                    styles,
                )

                slides.extend(split_slides)

    return slides


def _preprocess_blocks(blocks: list, max_table_rows: int) -> list:
    """큰 테이블 분할, 긴 문단 처리 등 전처리."""
    result = []
    for block in blocks:
        if block.get("type") == "table":
            rows = block.get("rows", [])
            if len(rows) > max_table_rows:
                # 테이블 분할
                chunks = _split_table(block, max_table_rows)
                result.extend(chunks)
            else:
                result.append(block)
        else:
            result.append(block)
    return result


def _split_table(table: dict, max_rows: int) -> list:
    """큰 테이블을 max_rows 단위로 분할. 헤더 반복."""
    headers = table["headers"]
    header_runs = table.get("header_runs", [])
    rows = table["rows"]
    row_runs = table.get("row_runs", [])

    chunks = []
    for start in range(0, len(rows), max_rows):
        end = min(start + max_rows, len(rows))
        chunk = {
            "type": "table",
            "headers": headers,
            "header_runs": header_runs,
            "rows": rows[start:end],
            "row_runs": row_runs[start:end] if row_runs else [],
        }
        chunks.append(chunk)
    return chunks


def _split_blocks_to_slides(blocks: list, title: str, subtitle: str,
                             styles: dict) -> list:
    """블록을 슬라이드 높이에 맞게 분할."""
    if not blocks:
        return [{
            "layout": "title_content",
            "title": title,
            "subtitle": subtitle,
            "blocks": [],
        }]

    content_height = int(DEFAULT_CONTENT_HEIGHT)
    content_width = int(DEFAULT_CONTENT_WIDTH)

    slides = []
    remaining = list(blocks)

    while remaining:
        fit_count = can_fit_blocks(remaining, styles, content_height, content_width)
        if fit_count <= 0:
            fit_count = 1  # 최소 1개는 넣기

        slide_blocks = remaining[:fit_count]
        remaining = remaining[fit_count:]

        # 테이블만 있는 슬라이드는 blank 레이아웃
        all_tables = all(b.get("type") == "table" for b in slide_blocks)
        # 이미지만 있는 슬라이드도 blank
        all_images = all(b.get("type") == "image" for b in slide_blocks)

        layout = "blank" if (all_tables or all_images) else "title_content"

        slides.append({
            "layout": layout,
            "title": title,
            "subtitle": subtitle,
            "blocks": slide_blocks,
        })

    return slides
