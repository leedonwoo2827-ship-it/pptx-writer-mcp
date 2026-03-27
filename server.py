# -*- coding: utf-8 -*-
"""
PPTX Writer MCP Server

Claude Desktop에서 사용하는 MCP 서버.
마크다운 텍스트 또는 MD 파일을 PPTX 프레젠테이션으로 변환합니다.

세션 기반 점진적 슬라이드 추가와 직접 변환 모드를 모두 지원합니다.

실행: python server.py
"""

import json
import os
import sys
import uuid
from pathlib import Path

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR / "src"))

from mcp.server.fastmcp import FastMCP
from md_parser import parse_slides_md, parse_markdown_tree
from slide_planner import plan_slides
from pptx_generator import PPTXGenerator

DEFAULT_STYLES_PATH = BASE_DIR / "pptx-styles.json"
DEFAULT_TEMPLATE_PATH = BASE_DIR / "templates" / "default.pptx"

mcp = FastMCP("pptx-writer")

# ---------------------------------------------------------------------------
# 세션 관리
# ---------------------------------------------------------------------------
_sessions: dict = {}  # session_id → {"generator": PPTXGenerator, "output_path": Path}


# ---------------------------------------------------------------------------
# 내부 유틸리티
# ---------------------------------------------------------------------------

def _resolve_styles_path(styles_file: str) -> Path:
    if not styles_file:
        return DEFAULT_STYLES_PATH
    p = Path(styles_file)
    return p if p.is_absolute() else BASE_DIR / p


def _load_styles(styles_path: Path) -> dict:
    if styles_path.exists():
        with open(styles_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _resolve_template_path(template_file: str) -> str:
    if not template_file:
        if DEFAULT_TEMPLATE_PATH.exists():
            return str(DEFAULT_TEMPLATE_PATH)
        return ""
    p = Path(template_file)
    if p.is_absolute():
        return str(p) if p.exists() else ""
    candidate = BASE_DIR / template_file
    return str(candidate) if candidate.exists() else ""


def _resolve_output_path(project_dir: str, output_file: str,
                         fallback_name: str = "output.pptx") -> Path:
    if project_dir:
        proj = Path(project_dir)
        out_dir = proj / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        name = Path(output_file).name if output_file else fallback_name
        return out_dir / name
    if output_file:
        p = Path(output_file)
        return p if p.is_absolute() else Path.home() / "Documents" / output_file
    return Path.home() / "Documents" / fallback_name


def _read_md_file(md_path: Path) -> str:
    raw = md_path.read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp949")


# ---------------------------------------------------------------------------
# MCP 도구: 세션 기반 (핵심)
# ---------------------------------------------------------------------------

@mcp.tool()
def create_pptx(
    output_file: str,
    template_file: str = "",
    styles_file: str = "",
    project_dir: str = "",
) -> str:
    """새 PPTX 프레젠테이션 세션을 시작합니다.

    대용량 문서를 슬라이드로 변환할 때 사용합니다.
    create_pptx → add_slides(반복) → finalize_pptx 순서로 호출하세요.

    Args:
        output_file: 저장할 PPTX 파일명 (예: proposal.pptx)
        template_file: 마스터 슬라이드 템플릿 경로 (생략 시 기본 템플릿)
        styles_file: 스타일 파일 경로 (생략 시 pptx-styles.json)
        project_dir: 프로젝트 폴더 경로. 지정 시 output/ 하위에 저장

    Returns:
        session_id (이후 add_slides, finalize_pptx에서 사용)
    """
    try:
        styles_path = _resolve_styles_path(styles_file)
        styles = _load_styles(styles_path)
        template_path = _resolve_template_path(template_file)
        out_path = _resolve_output_path(project_dir, output_file)
        base_dir = project_dir if project_dir else str(out_path.parent)

        generator = PPTXGenerator(
            template_path=template_path,
            styles=styles,
            base_dir=base_dir,
        )

        session_id = str(uuid.uuid4())[:8]
        _sessions[session_id] = {
            "generator": generator,
            "output_path": out_path,
            "styles": styles,
            "slide_count": 0,
        }

        template_name = Path(template_path).name if template_path else "기본 빈 템플릿"
        return (
            f"세션 생성 완료!\n"
            f"session_id: {session_id}\n"
            f"템플릿: {template_name}\n"
            f"출력 경로: {out_path}\n\n"
            f"이제 add_slides()로 슬라이드를 추가하세요.\n"
            f"슬라이드 마크다운 포맷:\n"
            f"---slide\n"
            f"layout: title_content\n"
            f"# 슬라이드 제목\n"
            f"## 부제목\n"
            f"- 불릿 항목\n"
            f"  - 서브 불릿\n"
            f"| 헤더 | 헤더 |\n"
            f"|------|------|\n"
            f"| 셀   | 셀   |\n\n"
            f"layout 옵션: section_header, title_content, blank, two_content, title_only"
        )

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def add_slides(
    session_id: str,
    slides_md: str,
) -> str:
    """현재 세션에 슬라이드를 추가합니다.

    Claude가 개조식으로 변환한 마크다운을 전달하면 PPTX 슬라이드로 렌더링합니다.

    슬라이드 마크다운 포맷:
    ---slide
    layout: section_header
    # I. 제안사 소개

    ---slide
    layout: title_content
    # 슬라이드 제목
    ## 부제목
    - 불릿 항목 1
    - 불릿 항목 2
      - 서브 불릿
    | 구분 | 내용 |
    |------|------|
    | 셀1  | 셀2  |

    ---slide
    layout: blank
    ![캡션](images/file.png)

    인라인 색상: {{red:텍스트}}, {{green:텍스트}}

    layout 옵션:
    - section_header: 섹션 구분 (H1 파트)
    - title_content: 제목+콘텐츠 (기본)
    - blank: 빈 슬라이드 (테이블/이미지)
    - two_content: 2단 레이아웃
    - title_only: 제목만

    Args:
        session_id: create_pptx에서 받은 세션 ID
        slides_md: 슬라이드 마크다운 텍스트 (---slide로 구분)

    Returns:
        추가된 슬라이드 수와 현재 총 슬라이드 수
    """
    if session_id not in _sessions:
        return f"오류: 유효하지 않은 세션 ID입니다: {session_id}\ncreate_pptx()로 새 세션을 시작하세요."

    try:
        session = _sessions[session_id]
        generator = session["generator"]

        # 슬라이드 마크다운 파싱
        slide_specs = parse_slides_md(slides_md)

        if not slide_specs:
            return "경고: 파싱된 슬라이드가 없습니다. ---slide 구분자를 확인하세요."

        # 슬라이드 추가
        added = generator.add_slides(slide_specs)
        session["slide_count"] = generator.slide_count

        return (
            f"슬라이드 추가 완료!\n"
            f"이번에 추가: {added}장\n"
            f"현재 총: {session['slide_count']}장\n\n"
            f"계속 add_slides()로 추가하거나 finalize_pptx()로 저장하세요."
        )

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def finalize_pptx(
    session_id: str,
) -> str:
    """PPTX 세션을 종료하고 파일을 저장합니다.

    Args:
        session_id: create_pptx에서 받은 세션 ID

    Returns:
        저장된 파일 경로, 크기, 슬라이드 수
    """
    if session_id not in _sessions:
        return f"오류: 유효하지 않은 세션 ID입니다: {session_id}"

    try:
        session = _sessions[session_id]
        generator = session["generator"]
        output_path = session["output_path"]

        # 저장
        generator.save(str(output_path))

        # 세션 정리
        slide_count = generator.slide_count
        del _sessions[session_id]

        if output_path.exists():
            size = output_path.stat().st_size
            return (
                f"PPTX 저장 완료!\n"
                f"파일: {output_path}\n"
                f"크기: {size:,} bytes\n"
                f"슬라이드: {slide_count}장"
            )
        return f"오류: 파일 생성에 실패했습니다: {output_path}"

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


# ---------------------------------------------------------------------------
# MCP 도구: 직접 변환 (규칙 기반 폴백)
# ---------------------------------------------------------------------------

@mcp.tool()
def convert_md_to_pptx(
    md_file: str,
    template_file: str = "",
    output_file: str = "",
    styles_file: str = "",
    project_dir: str = "",
) -> str:
    """마크다운(.md) 파일을 규칙 기반으로 PPTX 프레젠테이션으로 변환합니다.

    대용량 MD 파일을 자동으로 슬라이드로 분할합니다.
    H1 → 섹션 구분, H2 → 슬라이드 제목, H3+ → 콘텐츠로 매핑됩니다.

    Claude가 개조식으로 변환 후 세션 모드(create_pptx → add_slides → finalize_pptx)를
    사용하면 더 높은 품질의 결과를 얻을 수 있습니다.

    Args:
        md_file: 변환할 마크다운 파일 경로
        template_file: 마스터 슬라이드 템플릿 경로 (생략 시 기본 템플릿)
        output_file: 저장할 PPTX 파일명
        styles_file: 스타일 파일 경로 (생략 시 pptx-styles.json)
        project_dir: 프로젝트 폴더 경로. 지정 시 output/ 하위에 저장

    Returns:
        저장된 파일 경로, 크기, 슬라이드 수
    """
    md_path = Path(md_file)
    if not md_path.is_absolute():
        md_path = Path.home() / "Documents" / md_file
    if not md_path.exists():
        return f"오류: 파일을 찾을 수 없습니다: {md_path}"

    try:
        text_content = _read_md_file(md_path)
    except Exception as e:
        return f"오류: 파일 읽기 실패: {type(e).__name__}: {e}"

    fallback_name = md_path.with_suffix(".pptx").name
    out_path = _resolve_output_path(project_dir, output_file, fallback_name)

    try:
        styles_path = _resolve_styles_path(styles_file)
        styles = _load_styles(styles_path)
        template_path = _resolve_template_path(template_file)

        # 마크다운 트리 파싱
        tree = parse_markdown_tree(text_content)

        # 슬라이드 계획
        slide_specs = plan_slides(tree, styles)

        # PPTX 생성
        generator = PPTXGenerator(
            template_path=template_path,
            styles=styles,
            base_dir=str(md_path.parent),
        )
        generator.add_slides(slide_specs)
        generator.save(str(out_path))

        if out_path.exists():
            size = out_path.stat().st_size
            return (
                f"변환 완료!\n"
                f"원본: {md_path}\n"
                f"PPTX: {out_path} ({size:,} bytes)\n"
                f"슬라이드: {generator.slide_count}장"
            )
        return f"오류: 파일 생성에 실패했습니다: {out_path}"

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def convert_text_to_pptx(
    text_content: str,
    template_file: str = "",
    output_file: str = "",
    styles_file: str = "",
    project_dir: str = "",
) -> str:
    """마크다운 텍스트를 PPTX 프레젠테이션으로 변환합니다.

    슬라이드 마크다운 포맷(---slide 구분)을 감지하여 자동 처리합니다.
    일반 마크다운이면 규칙 기반으로 변환합니다.

    Args:
        text_content: 변환할 마크다운 텍스트
        template_file: 마스터 슬라이드 템플릿 경로
        output_file: 저장할 PPTX 파일명
        styles_file: 스타일 파일 경로
        project_dir: 프로젝트 폴더 경로

    Returns:
        저장된 파일 경로, 크기, 슬라이드 수
    """
    out_path = _resolve_output_path(project_dir, output_file)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        styles_path = _resolve_styles_path(styles_file)
        styles = _load_styles(styles_path)
        template_path = _resolve_template_path(template_file)

        # ---slide 포맷 감지
        if "---slide" in text_content:
            slide_specs = parse_slides_md(text_content)
        else:
            tree = parse_markdown_tree(text_content)
            slide_specs = plan_slides(tree, styles)

        generator = PPTXGenerator(
            template_path=template_path,
            styles=styles,
            base_dir=project_dir if project_dir else "",
        )
        generator.add_slides(slide_specs)
        generator.save(str(out_path))

        if out_path.exists():
            size = out_path.stat().st_size
            return (
                f"변환 완료!\n"
                f"PPTX: {out_path} ({size:,} bytes)\n"
                f"슬라이드: {generator.slide_count}장"
            )
        return f"오류: 파일 생성에 실패했습니다: {out_path}"

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


# ---------------------------------------------------------------------------
# MCP 도구: 유틸리티
# ---------------------------------------------------------------------------

@mcp.tool()
def get_pptx_styles(styles_file: str = "") -> str:
    """현재 pptx-styles.json의 스타일 설정을 반환합니다.

    Args:
        styles_file: 스타일 파일 경로 (생략 시 pptx-styles.json)

    Returns:
        현재 스타일 설정 JSON
    """
    styles_path = _resolve_styles_path(styles_file)
    if not styles_path.exists():
        return f"오류: 스타일 파일을 찾을 수 없습니다: {styles_path}"

    with open(styles_path, "r", encoding="utf-8") as f:
        styles = json.load(f)
    return f"스타일 파일: {styles_path}\n\n" + json.dumps(styles, ensure_ascii=False, indent=2)


@mcp.tool()
def get_template_info(template_file: str = "") -> str:
    """PPTX 템플릿의 슬라이드 레이아웃과 플레이스홀더 정보를 반환합니다.

    새 템플릿을 사용하기 전에 이 도구로 레이아웃을 확인하세요.

    Args:
        template_file: 템플릿 파일 경로 (생략 시 기본 템플릿)

    Returns:
        레이아웃 이름, 인덱스, 플레이스홀더 목록
    """
    template_path = _resolve_template_path(template_file)
    styles = _load_styles(_resolve_styles_path(""))

    try:
        generator = PPTXGenerator(template_path=template_path, styles=styles)
        info = generator.get_template_info()
        return json.dumps(info, ensure_ascii=False, indent=2)
    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


# ---------------------------------------------------------------------------
# 진입점
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
