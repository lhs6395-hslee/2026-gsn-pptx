---
name: ppt-generator
description: >
  Microsoft PowerPoint(.pptx) 파일을 자동으로 생성하는 스킬.
  사용자가 "PPT 만들어줘", "프레젠테이션 만들어줘", "발표 자료 만들어줘", 또는
  특정 주제와 함께 ppt/슬라이드/발표를 언급할 때 반드시 이 스킬을 사용한다.
  터미널에서 직접 주제를 인자로 넘겨도 동일하게 동작한다.
  AHE(Agentic Harness Engineering) 기법으로 하네스를 자동 개선한다.
---

# PPT 생성 스킬

## 핵심 원칙 — Python을 직접 실행하지 않는다

이 스킬은 bash 도구로 쉘 스크립트를 호출한다.
`python main.py` 같은 엔트리포인트를 실행하는 게 아니라,
Claude 자신이 오케스트레이터로서 각 단계를 직접 수행한다.

---

## 트리거 조건

다음 중 하나라도 해당하면 이 스킬을 즉시 사용한다:
- "PPT 만들어줘 / 만들어주세요"
- "프레젠테이션 만들어줘"
- "발표 자료 만들어줘"
- "슬라이드 만들어줘"
- `ppt <주제>` 형식의 터미널 입력
- 특정 주제 + 슬라이드 수 언급

---

## 필수 입력값 수집

스킬 실행 전 반드시 확인한다. 대화에 이미 있으면 다시 묻지 않는다.

| 항목 | 기본값 | 비고 |
|------|--------|------|
| 주제 | (필수) | |
| 슬라이드 수 | 10 | |
| 대상 청중 | 전문가 | |
| 템플릿 경로 | `~/.ppt-skill/templates/default.pptx` | |
| AHE 진화 여부 | OFF | 명시 요청 시에만 ON |

---

## 실행 흐름

### 0. 환경 확인

```bash
# 스킬 디렉토리 존재 확인
ls ~/.ppt-skill/harness/CLAUDE.md 2>/dev/null || {
  echo "스킬 미설치. setup.sh를 먼저 실행하세요."
  exit 1
}

# 템플릿 확인
TEMPLATE="${PPT_TEMPLATE:-$HOME/.ppt-skill/templates/default.pptx}"
ls "$TEMPLATE" || { echo "템플릿 없음: $TEMPLATE"; exit 1; }
```

### 1. 작업 디렉토리 생성

```bash
WORK="$HOME/.ppt-skill/runs/$(date +%Y%m%d_%H%M%S)_${TOPIC// /_}"
mkdir -p "$WORK"/{unpacked,runs,traces}
cp "$TEMPLATE" "$WORK/template.pptx"
cd "$WORK"
```

### 2. 하네스 컴포넌트 로드

```bash
# harness/ 디렉토리의 5개 파일을 읽어 작업 컨텍스트에 로드
SYSTEM_PROMPT=$(cat ~/.ppt-skill/harness/CLAUDE.md)
VERIFIER_RULES=$(cat ~/.ppt-skill/harness/verifier_rules.json)
MEMORY=$(cat ~/.ppt-skill/harness/long_term_memory.json)
```

### 3. 템플릿 분석

```bash
cd "$WORK"
python3 ~/.ppt-skill/scripts/office/unpack.py template.pptx unpacked/
python3 ~/.ppt-skill/scripts/thumbnail.py template.pptx thumbs --cols 3
extract-text template.pptx
```

분석 결과로 슬라이드 레이아웃 목록을 파악한다.

### 4. 슬라이드 계획 수립 (plan.json)

Claude가 직접 plan.json을 작성한다:
- 사용 가능한 레이아웃에서 각 슬라이드에 적합한 것을 선택
- 주제에 맞는 콘텐츠 구조 설계
- 전문가 수준의 내용 기획

```json
{
  "title": "발표 제목",
  "topic": "주제",
  "audience": "대상 청중",
  "slides": [
    {
      "index": 1,
      "template_file": "slide6.xml",
      "role": "cover",
      "title": "슬라이드 제목",
      "content": { ... }
    }
  ]
}
```

### 5. 슬라이드 XML 편집

각 슬라이드를 순서대로 편집한다.

**편집 전 미들웨어 Pre-hook:**
```bash
python3 -c "
import sys
from harness_hooks import pre_xml_edit
result = pre_xml_edit(open('$SLIDE_XML').read())
if not result['pass']:
    print('PRE-HOOK FAIL:', result['issues'])
    sys.exit(1)
"
```

**편집 후 Post-hook:**
```bash
python3 -c "
import xml.etree.ElementTree as ET
ET.parse('$SLIDE_XML')
print('XML valid')
"
```

### 6. 클린 & 패킹

```bash
python3 ~/.ppt-skill/scripts/clean.py unpacked/
python3 ~/.ppt-skill/scripts/office/pack.py \
  unpacked/ output.pptx --original template.pptx
```

### 7. 시각 QA

**soffice를 직접 호출 금지.** PowerPoint 우선 → LibreOffice 폴백 순서를 반드시 지킨다.

```bash
# ppt_generator.py의 visual_qa() 함수 사용 (PowerPoint AppleScript 우선)
python3 -c "
import sys, pathlib
sys.path.insert(0, '$HOME/Documents/claude/ppt_harness_project')
from ppt_generator import visual_qa
images = visual_qa(pathlib.Path('$WORK'), pathlib.Path('$WORK/output.pptx'))
print('QA images:', len(images), 'slides')
"
```

또는 Claude가 직접 AppleScript 호출:
```bash
osascript -e "
tell application \"Microsoft PowerPoint\"
    open (POSIX file \"$WORK/output.pptx\")
    delay 3
    save active presentation in (POSIX file \"$WORK/output.pdf\") as save as PDF
    delay 1
    close active presentation saving no
end tell
"
pdftoppm -jpeg -r 120 "$WORK/output.pdf" "$WORK/qa_images/slide"
```

> **왜 PowerPoint 우선인가**: 최종 파일(.pptx)은 PowerPoint에서 열어보는 것이 기준.
> LibreOffice는 Pretendard·한국어 폰트를 다르게 렌더링해 QA 결과가 실제와 다름.

슬라이드 이미지를 보고 확인:
- 텍스트 오버플로우
- 요소 겹침
- 플레이스홀더 잔여 (Lorem, 작성해주세요)
- 한국어 폰트 깨짐 (??? 표시)
- 헤더 중복

### 8. 콘텐츠 검증

```bash
extract-text output.pptx | grep -iE "lorem|작성해주세요|TODO|\[insert" && \
  echo "PLACEHOLDER FOUND" || echo "CONTENT CLEAN"
```

### 9. 결과 전달

```bash
# 최종 파일을 사용자 지정 위치로 복사
DEST="${OUTPUT_DIR:-$HOME/Desktop}/${TOPIC// /_}.pptx"
cp output.pptx "$DEST"
echo "완료: $DEST"
```

---

## AHE 진화 루프 (--evolve 시에만 실행)

진화 루프가 활성화된 경우, 위 5~8단계 후 추가로 실행:

### Trace 기록
각 슬라이드 결과를 `~/.ppt-skill/traces/` 에 JSON으로 저장한다.

### Digest 생성
```bash
python3 ~/.ppt-skill/ahe_tools/distill_digest.py \
  --run-id "$RUN_ID" \
  --traces-dir ~/.ppt-skill/traces/
```

### Evolve Agent 실행
digest를 읽고 `harness/` 파일들을 개선한다:
- 실패 패턴 → 해당 컴포넌트 수정
- 모든 편집에 예측 선언 (`change_manifest.json`)
- git commit으로 변경 추적

### 예측 검증
다음 라운드 결과와 이전 예측을 비교해 manifest를 업데이트한다.

---

## 오류 처리

| 상황 | 대응 |
|------|------|
| XML parse error | 해당 슬라이드 원본 복원 후 재편집 |
| 플레이스홀더 잔여 | 해당 슬라이드만 재편집 (최대 2회) |
| 시각 오버플로우 | 폰트 크기 축소 또는 내용 단축 |
| 패킹 실패 | clean.py 재실행 후 재시도 |
| 최대 2회 수정 후도 FAIL | best-effort 출력 + 사용자 알림 |

---

## 장기 기억 활용

`harness/long_term_memory.json`에서 다음을 읽어 활용한다:
- 특정 템플릿의 알려진 레이아웃 특성
- 이전 실행에서 발견한 수정 패턴
- 슬라이드 유형별 최적 접근법

AHE 루프가 실행될 때마다 이 파일이 자동으로 업데이트된다.
