---
description: 마크다운 파일을 규칙 기반으로 빠르게 PPTX로 변환한다. Claude 개조식 변환 없이 자동 분할.
argument-hint: "[마크다운 파일 경로]"
---

# /quick-convert

마크다운 파일을 규칙 기반으로 직접 PPTX로 변환한다.

## 실행 흐름

### Step 1: 변환

`convert_md_to_pptx` 도구를 호출한다.

- md_file: 마크다운 파일 경로
- 필요시 template_file, output_file, project_dir 지정

### Step 2: 결과 안내

생성된 PPTX 파일 경로, 크기, 슬라이드 수를 안내한다.
