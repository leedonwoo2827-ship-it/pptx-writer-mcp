# PPTX Writer MCP

마크다운을 대용량 PPTX 프레젠테이션으로 변환하는 Claude Desktop MCP 플러그인.

20만자 이상의 대용량 문서를 세션 기반으로 점진적 슬라이드 추가하여 PPTX로 변환합니다.
Claude가 줄글을 개조식으로 변환한 뒤 슬라이드를 생성하므로, 프레젠테이션에 적합한 형태로 자동 정리됩니다.

## 설치

### 방법 1: Claude Desktop Cowork 마켓플레이스
1. Claude Desktop → Cowork 탭 → 플러그인 탐색
2. `+` 버튼 → URL에 `your-github-id/pptx-writer-mcp` 입력 → 동기화

### 방법 2: 수동 설치
```bash
git clone https://github.com/your-github-id/pptx-writer-mcp.git
cd pptx-writer-mcp
install.bat
```

### 방법 3: 직접 설정
```bash
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
python create_template.py
```

`%APPDATA%\Claude\claude_desktop_config.json`에 추가:
```json
{
  "mcpServers": {
    "pptx-writer": {
      "command": "설치경로\\.venv\\Scripts\\python.exe",
      "args": ["설치경로\\server.py"]
    }
  }
}
```

## 사용법

### 세션 모드 (권장)

대용량 문서를 Claude가 섹션별로 읽고, 개조식으로 변환한 뒤 슬라이드를 추가합니다.

```
1. create_pptx(output_file="proposal.pptx")  → session_id 반환
2. add_slides(session_id, slides_md)          → 반복 호출
3. finalize_pptx(session_id)                  → 파일 저장
```

### 슬라이드 마크다운 포맷

`add_slides`에 전달하는 텍스트 포맷:

```markdown
---slide
layout: section_header
# I. 제안사 소개

---slide
layout: title_content
# 1. 일반현황
## 1.1 회사 소개
- 2000년 설립, 이러닝 전문기업
- KOICA ODA 사업 수행
  - 나이지리아, 아제르바이잔 등

---slide
layout: title_content
# 1. 일반현황
## 회사 일반현황
| 구분 | 내용 |
|------|------|
| 회사명 | factory |
| 대표자 | dekman |

---slide
layout: blank
![시스템 아키텍처](images/architecture.png)
```

### 레이아웃 옵션
| 레이아웃 | 용도 |
|----------|------|
| `section_header` | 파트/섹션 구분 제목 |
| `title_content` | 제목 + 불릿/텍스트 (기본) |
| `blank` | 테이블/이미지 전용 |
| `two_content` | 2단 레이아웃 |
| `title_only` | 제목만 |

### 인라인 마커
- `{{red:텍스트}}` → 빨간색
- `{{green:텍스트}}` → 초록색
- `**볼드**` → 굵은 글씨

### 직접 변환 모드
Claude 없이 규칙 기반으로 MD 파일을 자동 변환합니다:
```
convert_md_to_pptx(md_file="proposal.md")
```

## 회사 마스터 슬라이드 적용

설치 폴더의 `templates/default.pptx`를 회사 마스터 슬라이드 파일로 교체하면 됩니다.
또는 도구 호출 시 `template_file` 파라미터로 직접 경로를 지정할 수도 있습니다:
```
create_pptx(output_file="out.pptx", template_file="C:/경로/회사템플릿.pptx")
```

`get_template_info()`를 호출하면 템플릿의 레이아웃과 플레이스홀더를 확인할 수 있습니다.

## MCP 도구 목록

| 도구 | 설명 |
|------|------|
| `create_pptx` | 새 PPTX 세션 시작 |
| `add_slides` | 세션에 슬라이드 추가 |
| `finalize_pptx` | 세션 종료 및 파일 저장 |
| `convert_md_to_pptx` | MD 파일 직접 변환 (규칙 기반) |
| `convert_text_to_pptx` | 텍스트 직접 변환 |
| `get_pptx_styles` | 스타일 설정 조회 |
| `get_template_info` | 템플릿 레이아웃 정보 |

## 스타일 커스터마이징

`pptx-styles.json`을 편집하여 폰트, 크기, 색상, 레이아웃 매핑 등을 변경할 수 있습니다.
서버 재시작 없이 매 요청마다 자동으로 다시 읽힙니다.

## 시작 프롬프트

### 1. 대용량 MD → PPTX 변환 (세션 모드, 권장)

```
아래 마크다운 파일을 PPTX 프레젠테이션으로 만들어줘.
줄글은 개조식(불릿 포인트)으로 변환하고, 표는 그대로 유지해줘.
섹션별로 나눠서 작업해줘.

파일: C:/경로/proposal-body.md
출력: C:/경로/output/proposal.pptx
```

### 2. 회사 템플릿 적용

```
아래 파일을 우리 회사 마스터 슬라이드에 맞춰 PPTX로 변환해줘.

마크다운: C:/경로/proposal.md
회사 템플릿: C:/경로/company_master.pptx
출력: C:/경로/proposal_final.pptx

먼저 get_template_info로 템플릿 레이아웃을 확인하고 작업해줘.
```

### 3. 빠른 직접 변환 (규칙 기반)

```
이 마크다운 파일을 PPTX로 바로 변환해줘. 개조식 변환 없이 규칙 기반으로 빠르게.

convert_md_to_pptx(md_file="C:/경로/proposal.md")
```

### 4. 특정 섹션만 슬라이드로

```
아래 텍스트를 슬라이드로 만들어줘:

[여기에 텍스트 붙여넣기]

- 핵심 내용만 불릿으로 정리
- 표는 별도 슬라이드로
- section_header로 파트 구분해줘
```
