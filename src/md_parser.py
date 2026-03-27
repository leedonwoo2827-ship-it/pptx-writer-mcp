# -*- coding: utf-8 -*-
"""
마크다운 → 슬라이드 구조 파서

두 가지 모드 지원:
1. 슬라이드 마크다운 파싱 (---slide 구분자): Claude가 개조식으로 변환한 텍스트
2. 일반 마크다운 파싱 (H1-H5): 직접 변환용 (규칙 기반)

슬라이드 마크다운 포맷:
    ---slide
    layout: title_content
    # 슬라이드 제목
    ## 부제목
    - 불릿 항목 1
    - 불릿 항목 2
      - 서브 불릿
    | 헤더1 | 헤더2 |
    |-------|-------|
    | 셀1   | 셀2   |
    ![캡션](images/file.png)

인라인 컬러: {{red:텍스트}}, {{green:텍스트}}, {{blue:텍스트}}
인라인 볼드: **텍스트**
"""

import re
from typing import Optional


# ---------------------------------------------------------------------------
# 공개 API
# ---------------------------------------------------------------------------

def parse_slides_md(md_content: str) -> list:
    """슬라이드 마크다운(---slide 구분)을 파싱하여 SlideSpec 리스트 반환.

    Returns:
        [{"layout": "...", "title": "...", "subtitle": "...", "blocks": [...]}]
    """
    # ---slide 로 분할
    raw_slides = re.split(r'^---slide\s*$', md_content, flags=re.MULTILINE)

    slides = []
    for raw in raw_slides:
        raw = raw.strip()
        if not raw:
            continue
        slide = _parse_single_slide(raw)
        if slide:
            slides.append(slide)

    return slides


def parse_markdown_tree(md_content: str) -> dict:
    """일반 마크다운(H1-H5)을 계층 트리로 파싱. 직접변환 모드용.

    Returns:
        {"parts": [{"title": "...", "chapters": [{"title": "...", "sections": [...]}]}]}
    """
    lines = md_content.split("\n")
    parts = []
    current_part = None
    current_chapter = None
    current_section = None
    current_blocks = []

    def _flush_blocks():
        nonlocal current_blocks
        if current_blocks and current_section is not None:
            current_section["blocks"].extend(current_blocks)
        elif current_blocks and current_chapter is not None:
            if not current_chapter["sections"]:
                current_chapter["sections"].append({
                    "title": "",
                    "blocks": current_blocks,
                })
            else:
                current_chapter["sections"][-1]["blocks"].extend(current_blocks)
        current_blocks = []

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.rstrip()

        # H1
        if stripped.startswith("# ") and not stripped.startswith("## "):
            _flush_blocks()
            text = _clean_text(stripped[2:].strip())
            current_part = {"title": text, "chapters": []}
            current_chapter = None
            current_section = None
            parts.append(current_part)
            i += 1
            continue

        # H2
        if stripped.startswith("## ") and not stripped.startswith("### "):
            _flush_blocks()
            text = _clean_text(stripped[3:].strip())
            current_chapter = {"title": text, "sections": []}
            current_section = None
            if current_part is None:
                current_part = {"title": "", "chapters": []}
                parts.append(current_part)
            current_part["chapters"].append(current_chapter)
            i += 1
            continue

        # H3
        if stripped.startswith("### ") and not stripped.startswith("#### "):
            _flush_blocks()
            text = _clean_text(stripped[4:].strip())
            current_section = {"title": text, "blocks": []}
            if current_chapter is None:
                current_chapter = {"title": "", "sections": []}
                if current_part is None:
                    current_part = {"title": "", "chapters": []}
                    parts.append(current_part)
                current_part["chapters"].append(current_chapter)
            current_chapter["sections"].append(current_section)
            i += 1
            continue

        # H4
        if stripped.startswith("#### ") and not stripped.startswith("##### "):
            text = _clean_text(stripped[5:].strip())
            current_blocks.append({"type": "heading", "level": 4, "text": text})
            i += 1
            continue

        # H5
        if stripped.startswith("##### "):
            text = _clean_text(stripped[6:].strip())
            current_blocks.append({"type": "heading", "level": 5, "text": text})
            i += 1
            continue

        # Image
        img_match = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)', stripped)
        if img_match:
            current_blocks.append({
                "type": "image",
                "caption": img_match.group(1).strip(),
                "path": img_match.group(2).strip(),
            })
            i += 1
            continue

        # Table
        if stripped.startswith("|"):
            next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""
            if _is_separator_line(next_line):
                table, consumed = _parse_table(lines, i)
                if table:
                    current_blocks.append(table)
                    i += consumed
                    continue

        # Bullet list
        list_match = re.match(r"^(\s*)([-*+]|\d+\.)\s+(.+)", stripped)
        if list_match:
            indent = len(line) - len(line.lstrip())
            text = list_match.group(3).strip()
            level = 1 if indent == 0 else (2 if indent <= 4 else 3)
            current_blocks.append({
                "type": "bullet",
                "level": level,
                "runs": parse_inline_runs(text),
            })
            i += 1
            continue

        # Horizontal rule
        if re.match(r"^[-*_]{3,}\s*$", stripped):
            i += 1
            continue

        # Empty line
        if not stripped:
            i += 1
            continue

        # Paragraph
        runs = parse_inline_runs(stripped)
        current_blocks.append({"type": "paragraph", "runs": runs})
        i += 1

    _flush_blocks()
    return {"parts": parts}


def parse_inline_runs(text: str) -> list:
    """인라인 마커를 파싱하여 run 리스트 반환.

    {{red:텍스트}} → {"text": "텍스트", "color": "red"}
    **볼드** → {"text": "볼드", "bold": True}
    일반 → {"text": "일반"}
    """
    runs = []
    # 패턴: {{color:text}} 또는 **bold** 또는 일반 텍스트
    pattern = re.compile(
        r'\{\{(red|green|blue|yellow|black):([^}]+)\}\}'  # color marker
        r'|\*\*(.+?)\*\*'                                  # bold
        r'|([^{*]+|\{(?!\{)|(?:\*(?!\*)))'                # plain text
    )

    for m in pattern.finditer(text):
        if m.group(1):  # color marker
            runs.append({"text": m.group(2), "color": m.group(1)})
        elif m.group(3):  # bold
            runs.append({"text": m.group(3), "bold": True})
        elif m.group(4):  # plain
            if m.group(4).strip() or m.group(4) == " ":
                runs.append({"text": m.group(4)})

    # 빈 결과면 원본 텍스트를 그대로
    if not runs and text.strip():
        runs.append({"text": text})

    return runs


# ---------------------------------------------------------------------------
# 내부: 슬라이드 마크다운 파싱
# ---------------------------------------------------------------------------

def _parse_single_slide(raw: str) -> Optional[dict]:
    """하나의 ---slide 블록을 SlideSpec으로 파싱."""
    lines = raw.split("\n")
    spec = {
        "layout": "title_content",
        "title": "",
        "subtitle": "",
        "blocks": [],
    }

    i = 0
    # layout: 헤더 파싱
    while i < len(lines):
        line = lines[i].strip()
        layout_match = re.match(r'^layout:\s*(.+)', line)
        if layout_match:
            spec["layout"] = layout_match.group(1).strip()
            i += 1
            continue
        break

    # 나머지 콘텐츠 파싱
    while i < len(lines):
        line = lines[i]
        stripped = line.rstrip()

        # Title (#)
        if stripped.startswith("# ") and not stripped.startswith("## "):
            spec["title"] = _clean_text(stripped[2:].strip())
            i += 1
            continue

        # Subtitle (##)
        if stripped.startswith("## ") and not stripped.startswith("### "):
            spec["subtitle"] = _clean_text(stripped[3:].strip())
            i += 1
            continue

        # H3 heading → bold line
        if stripped.startswith("### "):
            text = stripped[4:].strip()
            spec["blocks"].append({
                "type": "heading",
                "level": 3,
                "runs": parse_inline_runs(text),
            })
            i += 1
            continue

        # Image
        img_match = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)', stripped)
        if img_match:
            spec["blocks"].append({
                "type": "image",
                "caption": img_match.group(1).strip(),
                "path": img_match.group(2).strip(),
            })
            i += 1
            continue

        # Table
        if stripped.startswith("|"):
            next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""
            if _is_separator_line(next_line):
                table, consumed = _parse_table(lines, i)
                if table:
                    spec["blocks"].append(table)
                    i += consumed
                    continue

        # Bullet
        list_match = re.match(r"^(\s*)([-*+]|\d+\.)\s+(.+)", stripped)
        if list_match:
            indent = len(line) - len(line.lstrip())
            text = list_match.group(3).strip()
            level = 1 if indent == 0 else (2 if indent <= 4 else 3)
            spec["blocks"].append({
                "type": "bullet",
                "level": level,
                "runs": parse_inline_runs(text),
            })
            i += 1
            continue

        # Empty line
        if not stripped:
            i += 1
            continue

        # Paragraph
        runs = parse_inline_runs(stripped)
        spec["blocks"].append({"type": "paragraph", "runs": runs})
        i += 1

    return spec if (spec["title"] or spec["blocks"]) else None


# ---------------------------------------------------------------------------
# 내부: 공통 헬퍼
# ---------------------------------------------------------------------------

def _clean_text(text: str) -> str:
    """볼드/컬러 마커를 제거하고 순수 텍스트 반환."""
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\{\{(?:red|green|blue|yellow|black):([^}]+)\}\}', r'\1', text)
    return text.strip()


def _is_separator_line(line: str) -> bool:
    stripped = line.strip()
    if not stripped:
        return False
    cleaned = stripped.replace("|", "").replace(":", "").replace("-", "").replace(" ", "")
    return len(cleaned) == 0 and "-" in stripped


def _parse_table(lines: list, start: int) -> tuple:
    """마크다운 테이블 파싱."""
    header_line = lines[start].strip()
    if not header_line.startswith("|"):
        return None, 0

    inner_header = header_line.strip("|")
    headers = [h.strip() for h in inner_header.split("|")]

    if not headers:
        return None, 0

    if start + 1 >= len(lines) or not _is_separator_line(lines[start + 1]):
        return None, 0

    col_count = len(headers)
    rows = []
    i = start + 2
    while i < len(lines):
        row_line = lines[i].strip()
        if not row_line.startswith("|"):
            break
        inner = row_line.strip("|")
        cells = [c.strip() for c in inner.split("|")]
        while len(cells) < col_count:
            cells.append("")
        cells = cells[:col_count]
        rows.append(cells)
        i += 1

    consumed = i - start
    # 각 셀의 인라인 마커를 run으로 파싱
    header_runs = [parse_inline_runs(h) for h in headers]
    row_runs = [[parse_inline_runs(cell) for cell in row] for row in rows]

    return {
        "type": "table",
        "headers": headers,
        "header_runs": header_runs,
        "rows": rows,
        "row_runs": row_runs,
    }, consumed
