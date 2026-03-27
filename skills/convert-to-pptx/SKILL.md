# /pptx-writer:convert-to-pptx

대용량 마크다운 파일을 PPTX 프레젠테이션으로 변환합니다.
줄글을 개조식(불릿 포인트)으로 변환하고, 표와 이미지는 그대로 유지합니다.

## 사용법

```
/pptx-writer:convert-to-pptx

파일: C:/경로/proposal-body.md
출력: proposal.pptx
project_dir: C:/경로/프로젝트폴더
```

## 실행 절차

### Step 1: 세션 시작
- `create_pptx` 도구를 호출하여 PPTX 세션을 시작합니다.
- `output_file`, `project_dir`, `template_file`(선택) 파라미터를 전달합니다.
- 반환된 `session_id`를 기억합니다.

### Step 2: 마크다운 파일 분석
- 마크다운 파일을 섹션별(H1, H2 단위)로 나눠 읽습니다.
- 전체 구조를 파악합니다.

### Step 3: 섹션별 슬라이드 변환 (반복)
- 각 섹션의 줄글을 **개조식(불릿 포인트)**으로 변환합니다.
- 아래 슬라이드 마크다운 포맷으로 작성합니다.
- `add_slides` 도구를 호출하여 슬라이드를 추가합니다.
- 모든 섹션을 처리할 때까지 반복합니다.

### Step 4: 저장
- `finalize_pptx` 도구를 호출하여 파일을 저장합니다.

## 슬라이드 마크다운 포맷

```markdown
---slide
layout: section_header
# 파트 제목

---slide
layout: title_content
# 슬라이드 제목
## 부제목
- 불릿 항목 1
- 불릿 항목 2
  - 서브 불릿

---slide
layout: title_content
# 제목
## 부제목
| 헤더1 | 헤더2 |
|-------|-------|
| 셀1   | 셀2   |

---slide
layout: blank
![캡션](images/파일.png)
```

## 레이아웃 옵션
- `section_header`: 파트/섹션 구분 제목
- `title_content`: 제목 + 불릿/텍스트 (기본)
- `blank`: 테이블/이미지 전용
- `two_content`: 2단 레이아웃
- `title_only`: 제목만

## 인라인 마커
- `{{red:텍스트}}` → 빨간색
- `{{green:텍스트}}` → 초록색
- `**볼드**` → 굵은 글씨

## 개조식 변환 원칙
- 서술형 문장을 핵심 키워드 + 수치 중심으로 압축
- 한 불릿에 1~2줄 이내
- 슬라이드당 불릿 5~7개 이내
- 표는 변환하지 않고 그대로 유지
- 이미지 참조도 그대로 유지
