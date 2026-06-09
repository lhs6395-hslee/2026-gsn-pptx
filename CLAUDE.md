# PPT AHE 하네스 프로젝트

> Claude Code에서 이 파일을 읽으면 이전 작업 컨텍스트를 그대로 이어받을 수 있다.

---

## ⛔ 필수 행동 규칙 (OVERRIDE — 모든 기본 동작보다 우선)

### 임의 데이터 생성 금지
- 파일 내용이 필요하면 반드시 `Read`로 직접 읽은 후 답변
- 명령 결과가 필요하면 `Bash`로 실제 실행 후 답변
- 수치(좌표, pt값, 색상코드, 슬라이드 번호)를 기억에서 생성 금지

### 할루시네이션 금지
- 불확실한 정보를 확실한 것처럼 단정 금지
- 라이브러리 API 사용 전 context7 MCP로 최신 문서 확인
- 코드 변경 전 반드시 해당 파일을 Read로 읽고 현재 상태 파악

### 시키지 않은 일 하지 않기 (OVERRIDE)
- 사용자가 요청하지 않은 수정·제안을 임의로 수행 금지
- **텍스트박스 2줄 줄바꿈은 문제가 아님** — 텍스트박스 cy 자동 확장으로 해결됨. 이를 "문제"로 언급하지 말 것
- 근거 없이 시각적 판단을 "문제"로 보고하지 말 것. 확인된 사실만 보고
- 이미 합의된 디자인 결정(맑은고딕 이미지설명, 파란색 강조색 등)을 다시 검토하거나 변경 제안하지 말 것

### 수정 범위 엄수 (OVERRIDE)
- "N페이지 수정", "이 파일 수정" → **정확히 그 대상만** 수정. 다른 페이지/파일에 동일 문제가 있어도 건드리지 않는다
- 동일 문제를 다른 곳에서 발견했을 때: 수정 전 반드시 "X페이지에도 같은 문제가 있습니다. 함께 수정할까요?"라고 먼저 확인
- 스크립트 작성 시 `for ... in all_slides` 같은 전체 순회 패턴은 요청 범위를 초과할 위험이 높음 — 대상을 명시적으로 제한할 것

### 잘못된 발언 기반 작업 즉시 되돌리기 (OVERRIDE)
- 잘못된 정보(틀린 페이지 번호, 잘못된 파일 식별 등)를 기반으로 수행한 작업은 오류 인지 즉시 되돌린다
- 수정 후 해당 작업이 잘못된 전제 위에서 수행됐음을 확인하면 → 사용자에게 알리고 revert 실행
- 예시: 틀린 페이지 번호로 slide10에 회색박스 적용 → 오류 확인 즉시 slide10 원복

### 수동 커스터마이징 재적용 의무
- `edit_slide` 또는 슬라이드 재생성 시, 이전에 수동으로 적용한 커스터마이징은 **반드시 함께 재적용**해야 함
- 재생성 전 수동 커스터마이징 목록을 파악하고 재적용 계획 수립
- 수동 커스터마이징이 포함된 슬라이드를 재생성할 때는 "이 슬라이드에 수동 적용된 요소가 있음"을 명시적으로 확인

### Python 수정 전 하네스 우선 원칙 (OVERRIDE)
- 어떤 변경이든 Python 코드를 수정하기 **전에** 반드시 자문한다: **"이 변경을 harness JSON으로 표현할 수 있는가?"**
- 가능하면 harness 먼저 수정한다. Python은 harness로 불가능한 경우에만 최후 수단으로 사용
- 판단 기준:
  - shape ID·역할·파라미터 변경 → `slide_shape_ids.json` 수정
  - 슬라이드 레이아웃 규칙 변경 → `layout_zone_map.json` 수정
  - 서식·색상·포맷 변경 → `common_formatting.json` 수정
  - placeholder 탐지 패턴 변경 → `placeholder_patterns.json` 수정
  - 새 슬라이드 동작 규칙 → `long_term_memory.json` 수정
- Python 수정이 불가피한 경우에도 수정 범위를 최소화하고, 파라미터는 하네스에서 읽도록 설계

### 슬라이드 편집 전 템플릿·하네스 확인 의무 (OVERRIDE)
- 슬라이드를 생성하거나 수정하기 전 **반드시** 아래를 확인한다
  1. `harness/long_term_memory.json` → `slide_layout_hints` + 해당 슬라이드의 `known_failure_fixes_*` 항목
  2. `harness/layout_zone_map.json` → 해당 template_file의 zone 구조 (어떤 shape에 무엇이 들어가는지)
  3. `harness/slide_shape_ids.json` → 해당 슬라이드의 shape ID · 역할 · truncate 파라미터
- **템플릿을 보지 않고 추측으로 zone 역할 판단 금지**
  - "이 shape이 이미지 슬롯인지 설명 슬롯인지"는 반드시 zone_map과 long_term_memory 근거로 판단
  - XML의 noFill·텍스트 여부만으로 역할 추론 금지
- 타임라인/흐름 레이아웃(slide30 등) 편집 시 `known_failure_fixes.flow_card_right_cutoff_slide5` 제약 적용 필수 (v2~v6는 단일 dict로 통합됨)
  - 카드 우측 경계: card.left + card.width ≤ slide_width - 200000
  - 카드 하단 경계: card.top + card.height ≤ slide_height - 200000

### 답변 신뢰도 태그 의무화
모든 실질적 정보 제공 시 아래 태그를 붙인다. 순수 절차 설명·실행 결과 직접 인용은 생략 가능.

| 태그 | 의미 | 사용 조건 |
|------|------|---------|
| `[공식]` | 직접 확인된 사실 | 방금 Read/Bash로 확인, 공식 문서 인용 |
| `[추측]` | 합리적 추론 | 파일 미확인, 맥락 기반 유추 |
| `[미확인]` | 출처 불분명 | 다른 세션 전달, 기억 기반, 검증 전 |

---

## 프로젝트 목표

Microsoft PowerPoint(.pptx)를 **주제만 입력하면 자동 생성**하는 시스템.
AHE(Agentic Harness Engineering, arXiv:2604.25850) 기법으로 하네스가 실행마다 스스로 개선된다.

- Python 파일은 설치 시 1회만 생성. 주제가 바뀌어도 코드는 그대로.
- Claude Code 스킬로 동작 — `python main.py`를 직접 실행하지 않음.
- 실행마다 `~/.ppt-skill/runs/<날짜>_<주제>/` 작업 폴더만 새로 생성.

원칙은 `harness/AHE_PRINCIPLES.md`에 명문화(2026-06 공식 출처 기반). 모든 하네스 변경은 이 문서 기준으로 판단.

| AHE 기둥 | 구현 위치 | 역할 |
|---------|-----------|------|
| ❶ Component Observability | `harness/` 파일들 | 하네스를 파일로 분리, git 추적 |
| ❷ Experience Observability | `_record_run_experience` → `long_term_memory.json`(runs/success_rate) · `evolution/last_run_digest.json` (헤드리스 `--evolve`는 `ahe_loop.distill_digest`) | 매 실행 경험 자동 기록 |
| ❸ Decision Observability | `harness/change_manifest.jsonl` (falsifiable-contract 원장) | 편집+예측 선언 → 다음 라운드 검증 |
