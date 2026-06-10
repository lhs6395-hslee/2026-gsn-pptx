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
| 템플릿 경로 | `./template/2026_PPT Template.pptx` | 프로젝트 디렉토리 내 |
| 작업 디렉토리 | `./result/tmp/` | 임시 파일 생성 위치 |
| AHE 진화 여부 | OFF | 명시 요청 시에만 ON |

---

## 실행 흐름

### 주 경로 — 통합 엔진 호출 (권장)

생성·편집·패킹·시각 QA 로직은 모두 `ppt_generator.py` 엔진에 구현되어 있다.
슬라이드 XML을 손으로 편집하지 말고 **검증된 엔진을 호출**한다 (재현성·신뢰성 확보).

**중요**: 프로젝트 디렉토리에서 실행해야 한다.

```bash
cd /Users/toule/Documents/claude/ppt-skill  # 프로젝트 디렉토리로 이동

python3 - <<'PY'
import sys, pathlib
from datetime import datetime

# 프로젝트 디렉토리
PROJECT_DIR = pathlib.Path.cwd()
sys.path.insert(0, str(PROJECT_DIR))

from ppt_generator import run_ppt_generation

topic = "<주제>"
work_dir = PROJECT_DIR / "result" / "tmp" / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{topic.replace(' ', '_')}"
template_path = PROJECT_DIR / "template" / "2026_PPT Template.pptx"

out, vision_issues = run_ppt_generation(
    topic=topic,
    template_path=template_path,
    work_dir=work_dir,
    audience="<청중>",
    n_slides=<장수>,
    inline_vision_qa=False,   # AHE_PRINCIPLES §2: 엔진 자기-QA 끔 → 11단계 독립 QA 에이전트가 검증(확증편향 차단)
    # layout_from_pptx=PROJECT_DIR / "reference.pptx",  # 레이아웃 고정 시 지정
)
print("완료:", out, "| vision_issues:", vision_issues)
PY
```

엔진(결정적 hands)이 자동으로 수행하는 것:
- 템플릿 언팩 → 41슬라이드 존 분석(`layout_zone_map.json`)
- plan 기반 존 맵 슬라이드 편집(아이콘/이미지 자리·정렬 정확) → 패킹 → 결정적 검증(verifier_rules·placeholder·폰트)
- 매 실행 후 경험 자동 기록(`_record_run_experience` → `long_term_memory.json`·`evolution/last_run_digest.json`)

> **plan과 QA는 brain 작업**이다(AHE_PRINCIPLES §6 brain/hands 분리). plan은 오케스트레이터(이 Claude 세션)가
> 작성하고(5단계), QA는 **생성 컨텍스트를 모르는 독립 격리 에이전트**가 수행한다(11단계, §2 확증편향 차단).
> `inline_vision_qa=False`로 엔진의 인라인 자기-검증을 끈다. (headless `python main.py` 폴백만 인라인 vision 사용)

생성된 `output.pptx`를 사용자 지정 위치로 복사해 전달한다.
아래 0~9단계는 엔진 내부 동작의 참고용 상세 설명이다.

---

### 0. 환경 확인

```bash
# 프로젝트 디렉토리 확인
cd /Users/toule/Documents/claude/ppt-skill
ls harness/CLAUDE.md 2>/dev/null || {
  echo "스킬 미설치. setup.sh를 먼저 실행하세요."
  exit 1
}

# 템플릿 확인
TEMPLATE="./template/2026_PPT Template.pptx"
ls "$TEMPLATE" || { echo "템플릿 없음: $TEMPLATE"; exit 1; }
```

### 1. 작업 디렉토리 생성

```bash
WORK="./result/tmp/$(date +%Y%m%d_%H%M%S)_${TOPIC// /_}"
mkdir -p "$WORK"/{unpacked,traces}
cp "$TEMPLATE" "$WORK/template.pptx"
cd "$WORK"
```

### 2. 하네스 컴포넌트 로드

```bash
# harness/ 디렉토리의 파일들을 읽어 작업 컨텍스트에 로드
SYSTEM_PROMPT=$(cat ../harness/CLAUDE.md)
VERIFIER_RULES=$(cat ../harness/verifier_rules.json)
ZONE_MAP=$(cat ../harness/layout_zone_map.json)
```

### 3. 템플릿 분석

```bash
cd "$WORK"
python3 ~/.ppt-skill/scripts/office/unpack.py template.pptx unpacked/
python3 ~/.ppt-skill/scripts/thumbnail.py template.pptx thumbs --cols 3
extract-text template.pptx
```

분석 결과로 슬라이드 레이아웃 목록을 파악한다.

### 4. 레이아웃 선택 — 반드시 이 순서대로

1. `~/.ppt-skill/harness/layout_features.json` Read → 알고리즘 필터 (신뢰도·컬럼수·아이콘수 매칭)
2. `~/.ppt-skill/harness/thumbnails/slide{NN}.png` Read → 시각 확인 (이미지 로드)
3. `~/.ppt-skill/harness/long_term_memory.json` Read → 과거 성공/실패 사례 참조

```
아키텍처/파이프라인 매핑 → slide38
기능/특징 3가지          → slide13 또는 slide15
기능/특징 4가지          → slide14 또는 slide16
4단계 프로세스           → slide30
로드맵 연도별            → slide29
현황→목표 비교           → slide36 (개념) / slide37 (수치)
이관/전환 효과           → slide35
표지/목차/섹션/QA/감사   → slide6/7/8/44/46
```

### 5. 슬라이드 계획 수립 (plan.json)

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

### 6. 슬라이드 XML 편집

각 슬라이드를 순서대로 편집한다.

**편집 전/후 검사는 엔진이 내부 처리한다** — 별도 훅 import 불필요:
- 이스케이프·lang=ko-KR·endParaRPr 순서: `apply_known_fixes`(편집 전) + `_write_xml`(편집 시)
- XML 유효성: `pack_output`이 패킹 전 `ET.parse`로 전 슬라이드 검증

(수동 점검이 필요하면 `python3 -c "import xml.etree.ElementTree as ET; ET.parse('$SLIDE_XML')"`로 XML 유효성만 확인.)

### 7. 클린 & 패킹

```bash
python3 ~/.ppt-skill/scripts/clean.py unpacked/
python3 ~/.ppt-skill/scripts/office/pack.py \
  unpacked/ output.pptx --original template.pptx
```

### 8. 시각 QA

**PowerPoint만 사용. LibreOffice(soffice) 호출 완전 금지.**

```bash
# ppt_generator.py의 visual_qa() 함수 사용 (PowerPoint AppleScript)
python3 -c "
import sys, pathlib
sys.path.insert(0, '$HOME/.ppt-skill')
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

> **PowerPoint만 사용하는 이유**: 최종 파일(.pptx)은 PowerPoint 렌더링이 기준.
> LibreOffice는 Pretendard·한국어 폰트를 다르게 렌더링해 QA 결과가 실제와 다름 — 사용 금지.

슬라이드 이미지를 보고 확인:
- 텍스트 오버플로우
- 요소 겹침
- 플레이스홀더 잔여 (Lorem, 작성해주세요)
- 한국어 폰트 깨짐 (??? 표시)
- 헤더 중복

### 9. 콘텐츠 검증

```bash
extract-text output.pptx | grep -iE "lorem|작성해주세요|TODO|\[insert" && \
  echo "PLACEHOLDER FOUND" || echo "CONTENT CLEAN"
```

### 9.5. 서식 QA (format_qa.py)

텍스트·폰트 서식을 원본 템플릿과 비교 검증한다. 생성 완료 후 선택적으로 실행.

```bash
python3 scripts/format_qa.py output.pptx template/2026_PPT\ Template.pptx
```

- 출력: 슬라이드별 서식 차이 리포트 (폰트명·크기·색상 불일치)
- 지오메트리 비교: shape 위치·크기 허용 오차(±5%) 초과 시 경고

### 10. 결과 전달

```bash
# 최종 파일을 사용자 지정 위치로 복사
DEST="${OUTPUT_DIR:-$HOME/Desktop}/${TOPIC// /_}.pptx"
cp output.pptx "$DEST"
echo "완료: $DEST"
```

### 11. 자동 QA (생성 완료 후 즉시 실행 — 사용자 요청 불필요)

**생성이 완료되면 반드시 이 단계를 자동으로 실행한다. 사용자가 별도로 `/qa_review`를 호출하지 않아도 된다.**

생성된 PPTX 경로(`DEST` 또는 `run_ppt_generation()` 반환값 `out`)를 확보한 후,
**Agent tool을 사용해 독립 에이전트를 spawn**한다.
생성 컨텍스트 없이 처음 보는 리뷰어 관점으로 QA를 수행하도록 아래 프롬프트를 전달한다.

```
당신은 PPT QA 리뷰어입니다. 생성 컨텍스트 없이 독립적으로 아래 PPTX를 검증하세요.

PPTX 경로: <생성된 PPTX 절대 경로>
하네스 경로: /Users/toule/Documents/claude/ppt-skill/harness/

## 단계 1: 텍스트/XML 검증

1. PPTX 언팩 (임시 폴더는 프로젝트 내 result/tmp 사용 — /tmp 금지):
   mkdir -p /Users/toule/Documents/claude/ppt-skill/result/tmp/qa_auto && unzip -o "<생성된 PPTX 절대 경로>" -d /Users/toule/Documents/claude/ppt-skill/result/tmp/qa_auto/

2. placeholder_patterns.json 읽기:
   /Users/toule/Documents/claude/ppt-skill/harness/placeholder_patterns.json

3. 모든 슬라이드 XML에서 잔류 placeholder 탐지 (python3 인라인):
   python3 -c "
   import json, re, os, glob
   from xml.etree import ElementTree as ET

   with open('/Users/toule/Documents/claude/ppt-skill/harness/placeholder_patterns.json') as f:
       patterns = json.load(f)['patterns']

   ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
   issues = []

   for xml_path in sorted(glob.glob('/Users/toule/Documents/claude/ppt-skill/result/tmp/qa_auto/ppt/slides/slide[0-9]*.xml')):
       slide_name = os.path.basename(xml_path)
       try:
           tree = ET.parse(xml_path)
           root = tree.getroot()
           for sp in root.iter():
               tag = sp.tag
               if not tag.endswith('}sp'):
                   continue
               nvSpPr = sp.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}nvSpPr') or \
                        sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}nvSpPr')
               cNvPr = sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr')
               if cNvPr is None:
                   for child in sp.iter():
                       if child.tag.endswith('}cNvPr'):
                           cNvPr = child
                           break
               shape_id = cNvPr.get('id') if cNvPr is not None else '?'
               texts = [t.text or '' for t in sp.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t')]
               full_text = ''.join(texts).strip()
               if not full_text:
                   continue
               for pat in patterns:
                   if re.search(pat, full_text, re.IGNORECASE | re.MULTILINE):
                       issues.append({'slide': slide_name, 'shape_id': shape_id, 'pattern': pat, 'text': full_text[:80]})
                       break
       except Exception as e:
           issues.append({'slide': slide_name, 'shape_id': 'parse_error', 'pattern': str(e), 'text': ''})

   print(json.dumps(issues, ensure_ascii=False, indent=2))
   "

4. 결과를 정리하여 잔류 placeholder 목록 출력

## 단계 2: 가시성 검증 (PowerPoint PDF export)

1. PDF 저장 경로: `/Users/toule/Documents/claude/ppt-skill/result/tmp/qa_review.pdf` (항상 고정 경로 사용 — 프로젝트 폴더에 저장 금지)

2. osascript로 PowerPoint PDF export:
   osascript << 'APPLESCRIPT'
   tell application "Microsoft PowerPoint"
       activate
       set pptFile to POSIX file "<생성된 PPTX 절대 경로>"
       set pdfFile to POSIX file "/Users/toule/Documents/claude/ppt-skill/result/tmp/qa_review.pdf"
       open pptFile
       delay 3
       set theDoc to active presentation
       save theDoc in pdfFile as save as PDF
       delay 2
       close theDoc saving no
   end tell
   APPLESCRIPT

3. PDF 생성 확인: ls -la /Users/toule/Documents/claude/ppt-skill/result/tmp/qa_review.pdf

   **⚠️ QA 진행 중 PDF 삭제 금지**: QA 리포트를 출력하는 동안 PDF를 삭제하지 않는다. 사용자가 리포트와 함께 PDF를 직접 열어 확인할 수 있어야 한다. PDF 삭제는 Step 10(result 폴더 복사) 완료 후에만 허용한다.

4. PDF 모든 페이지를 Read tool로 읽어 슬라이드별 가시성 확인:
   - placeholder 텍스트 잔류 여부
   - 텍스트가 슬라이드 밖으로 넘침 여부
   - 비어있어야 할 영역에 텍스트가 있는지
   - 내용이 있어야 할 영역이 비어있는지

## 결과 리포트 형식

아래 구조로 리포트를 반환하세요:

### 텍스트/XML 검증 결과
| 슬라이드 | Shape ID | 패턴 | 텍스트 |
|---|---|---|---|
| ... | ... | ... | ... |

### 가시성 검증 결과
| 페이지 | 레이아웃 | 상태 | 이슈 |
|---|---|---|---|
| ... | ... | ... | ... |

### 종합 판정
- 통과: 이슈 없는 슬라이드 목록
- 수정 필요: 이슈 있는 슬라이드 + 이슈 요약
```

에이전트 결과를 받아 통합 리포트를 출력하고,
수정 필요 항목이 있으면 사용자에게 수정 여부를 확인 후 진행한다.

> **확증편향 차단(AHE_PRINCIPLES §2)**: 이 QA 에이전트는 plan·생성 근거를 전달받지 않는다.
> 결과 PPTX·렌더 이미지·원 주제만 보고 판정한다. 오케스트레이터(생성한 세션)가 직접 자기 결과를
> 평가하지 말 것 — 반드시 독립 에이전트의 판정을 받는다.

### 11.5 AHE 피드백 — QA 판정 기록 + 발견을 하네스에 반영 (falsifiable-contract + 통보)

**(0) 독립 QA 판정을 경험 기록에 반영 (필수, F1 루프 닫기)** — skill 경로는 엔진이 QA를 안 해
`qa_ok=None`(보류)으로 기록돼 있다. 독립 QA 에이전트의 종합 판정(통과/수정필요)을 확정값으로 채운다.
**off-by-one 회피(#10): 생성 직후 `evolution/last_run_digest.json`의 `run_id`를 읽어 그 run을 정확히 닫는다**
(생성과 QA 사이에 다른 생성이 끼면 `runs[-1]`이 엉뚱한 run을 가리킬 수 있음):
```bash
RUN_ID=$(python3 -c "import json; print(json.load(open('evolution/last_run_digest.json')).get('run_id',''))")
python3 -c "import sys; sys.path.insert(0,'.'); from ppt_generator import update_last_run_qa; update_last_run_qa($QA_OK, run_id='$RUN_ID' or None)"
# $QA_OK = True(이슈 없음) / False(수정 필요). success_rate가 독립 QA 결과를 반영하게 됨.
# run_id를 못 구하면(레거시 digest) update_last_run_qa($QA_OK) 만 호출 → qa_ok=None인 가장 최근 보류 run을 닫는다.
```

**(0.5) QA 결함 → 진화 루프 환류 (#12 evolve-manual-trigger-gap, human-in-loop 게이트 보존)** —
이전에는 self-healing이 `--evolve`(헤드리스) 또는 인라인 vision 이슈에만 의존해, skill 경로(엔진이
vision을 안 돌려 `vision_issues=0` 고정 + main.py 미경유)에서 독립 QA가 결함을 찾아도 `run_evolve_loop`에
도달할 수 없었다. 이제 `ahe_loop.maybe_run_evolve_loop`이 그 환류 경로를 제공한다.
독립 QA가 `qa_ok=False`(수정 필요)를 기록했다면, 위 (0) 직후 다음을 호출해 **진화를 제안**한다:
```bash
# qa_ok=False일 때만 의미 있음. approved 기본 False → 자동 실행하지 않고 제안만 출력한다.
python3 -c "import sys; sys.path.insert(0,'.'); from pathlib import Path; from ahe_loop import maybe_run_evolve_loop; \
maybe_run_evolve_loop(Path('$WORK_DIR'), '$TOPIC', qa_ok=False)"
```
> **human-in-loop §1 (절대 보존)**: `maybe_run_evolve_loop`은 `approved=True`(또는 `explicit=True`,
> 즉 사람이 직접 `--evolve` 지정) 없이는 **절대 `run_evolve_loop`를 실행하지 않는다**. 자동 머지·자동
> 진화를 강제하지 않는다. 결함이 감지되면 오케스트레이터는 사용자에게 진화 실행 여부를 물어보고,
> **사용자가 승인한 경우에만** `maybe_run_evolve_loop(..., approved=True)` 또는 `run_evolve_loop`를 직접 호출한다.
> 승인 없이는 아래 (1)~(4)의 수동 하네스 반영 절차를 따른다(둘 중 하나만 수행 — 진화 루프와 수동 반영은 중복).

그 다음, 독립 QA가 **반복적·체계적 결함**(특정 레이아웃/존이 또 깨짐, placeholder 잔류 패턴 등)을 발견하면,
일회성 수정에 그치지 말고 **하네스 변경**으로 반영한다. 절차:

1. **stale 가정 자문(§3)**: 이 결함이 "약한 모델용 scaffolding이 빠져서"인가, 아니면 "하네스 데이터·로직 버그"인가?
   - 데이터·로직 버그 → 해당 `harness/*.json` 또는 편집기 수정 (하네스 우선).
   - 현 모델이면 충분한 일회성 콘텐츠 문제 → 하네스 변경 없이 콘텐츠만 교정.
2. **falsifiable-contract 기록**: 하네스를 바꿨다면 `harness/change_manifest.jsonl`에 1줄 append —
   `{date, id, change, files, evidence, root_cause, fix, predicted_fixes, regression_risk, verification:"pending"}`.
3. **사용자 통보**: 변경 건마다 manifest 항목을 요약해 사용자에게 알린다 (필수).
4. **다음 실행 검증**: 다음 생성에서 예측이 맞았는지(`predicted_fixes` 해결 / `regression_risk` 미발생) 확인하고
   manifest의 `verification`을 `verified`/`refuted`로 갱신한다.

> 엔진은 매 실행 후 `_record_run_experience`로 `long_term_memory.json`(runs/success_rate)과
> `evolution/last_run_digest.json`을 자동 갱신한다(경험 관찰성 ❷ + auto-update). 하네스 *변경* 기록은
> 위 manifest(결정 관찰성 ❸)가 담당한다 — 둘은 역할이 다르다.

---

## AHE 진화 루프 (--evolve 시에만 실행)

진화 루프가 활성화된 경우, 위 6~9단계 후 추가로 실행:

### Trace 기록
각 슬라이드 결과를 `~/.ppt-skill/traces/` 에 JSON으로 저장한다.

### Digest 생성
`--evolve` 헤드리스 경로는 `ahe_loop.run_evolve_loop`(main.py)가 자체 `distill_digest`로 트레이스를 압축한다.
일반(skill) 경로는 매 실행 후 `_record_run_experience`가 `long_term_memory.json`(runs/success_rate)과
`evolution/last_run_digest.json`을 자동 기록한다(경험 관찰성 ❷).

### Evolve Agent 실행
digest를 읽고 `harness/` 파일들을 개선한다:
- 실패 패턴 → 해당 컴포넌트 수정 (하네스 우선)
- 모든 편집에 예측 선언 → `harness/change_manifest.jsonl` (falsifiable-contract, AHE_PRINCIPLES §4)
- git commit으로 변경 추적

### 예측 검증
다음 라운드 결과와 이전 예측을 비교해 `change_manifest.jsonl`의 `verification`을 verified/refuted로 갱신한다.

---

## 알려진 이슈 & 해결법

1. **목차(slide7) 번호 정렬**: `anchor="ctr"` + `spcPct val="150000"` 필수
2. **섹션(slide8) 서브아이템 정렬**: `anchorCtr="0"` 추가 필요
3. **헤더 중복**: `>PPT <` 교체 후 다음 run 중복 발생 → regex로 제거
4. **XML 이스케이프**: `&` → `&amp;`, `<` → `&lt;` 필수. 미처리 시 pack 실패
5. **lorem ipsum 잔여**: 여러 `<a:r>` run에 분산됨. 단순 replace 아닌 regex sub 필요

---

## 오류 처리

| 상황 | 대응 |
|------|------|
| XML parse error | 해당 슬라이드 원본 복원 후 재편집 |
| 플레이스홀더 잔여 | 해당 슬라이드만 재편집 (최대 2회) |
| 시각 오버플로우 | 폰트 크기 축소 또는 내용 단축 |
| 패킹 실패 | clean.py 재실행 후 재시도 |
| 최대 2회 수정 후도 FAIL | best-effort 출력 + 사용자 알림 |
| 특정 슬라이드만 재빌드 필요 | `python3 scripts/rebuild_slides.py <작업디렉토리> --slides <번호목록>` |

---

## 장기 기억 활용

`harness/long_term_memory.json`에서 다음을 읽어 활용한다:
- 특정 템플릿의 알려진 레이아웃 특성
- 이전 실행에서 발견한 수정 패턴
- 슬라이드 유형별 최적 접근법

AHE 루프가 실행될 때마다 이 파일이 자동으로 업데이트된다.
