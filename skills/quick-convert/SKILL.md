# /pptx-writer:quick-convert

마크다운 파일을 규칙 기반으로 빠르게 PPTX로 변환합니다.
Claude가 개조식 변환을 하지 않고, 자동 규칙으로 슬라이드를 분할합니다.

## 사용법

```
/pptx-writer:quick-convert

파일: C:/경로/proposal.md
```

## 실행 절차

### Step 1: 변환
- `convert_md_to_pptx` 도구를 호출합니다.
- `md_file` 파라미터에 마크다운 파일 경로를 전달합니다.
- 필요시 `template_file`, `output_file`, `project_dir` 파라미터를 추가합니다.

### Step 2: 결과 확인
- 생성된 PPTX 파일 경로, 크기, 슬라이드 수를 사용자에게 안내합니다.
