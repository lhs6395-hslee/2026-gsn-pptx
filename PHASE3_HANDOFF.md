<!-- 핸드오프 문서: 맥(분석/검증) → 윈도우(실행) -->
<!-- 이 문서는 윈도우 네이티브 Claude Code 세션이 Phase 3(실제 코드 수정)를 수행하기 위한 자기완결 지침이다. -->

# Phase 3 실행 핸드오프 (윈도우 네이티브 세션용)

## 0. 컨텍스트 (이 작업이 무엇인가)

`2026-gsn-pptx` 코드베이스(핵심 `ppt_generator.py` ~7,000 LOC 모놀리스)의 리팩토링이다.
**분석(Phase 1)과 적대적 검증(Phase 2)은 맥에서 이미 완료**됐고, 그 산출물이 이 브랜치에 있다:

- `REFACTOR_DIAGNOSE_REPORT.md` — 사람이 읽는 진단 리포트 (confirmed 19 / needs-care 12 / rejected 3)
- `REFACTOR_DIAGNOSE_RAW.json` — 기계 참조용 후보별 원시 데이터(검증 verdict 포함)

**너의 임무 = Phase 3 (실제 코드 수정).** 위 리포트의 §6 "실행 권장 순서"를 그대로 따른다.

## 1. 시작 전 필수 (분석/검증 재실행 금지)

```
git checkout refactor/2026-05-modernization
git pull
```
- `REFACTOR_DIAGNOSE_REPORT.md` 전체와 `REFACTOR_DIAGNOSE_RAW.json`을 먼저 읽는다.
- 이 repo의 `CLAUDE.md` + `SKILL.md` + `harness/AHE_PRINCIPLES.md`를 읽고 하네스 우선 원칙·셀프힐링 규약을 준수한다.
- **분석을 다시 하지 마라.** 검증까지 끝난 결론이다. 의심되면 해당 후보의 `REFACTOR_DIAGNOSE_RAW.json` verdict/reasoning을 근거로만 재확인한다.

## 2. 모델 전략 (하이브리드)

- **기본 = Sonnet** (claude-sonnet-4-6): 코드 작성, 삭제, config 이전, 테스트 작성 등 패턴형 작업.
- **사고형 step만 Opus** (claude-opus-4-8): AHE 관찰성/자가치유 재설계(Phase 3-4 #10~13), needs-care 중 동작 분기 보존이 까다로운 것(#14 dedupe, #18 flow, #19 fold).
- Workflow 오케스트레이션 시 `agent(..., {model})`로 step별 배정. 윈도우 PC 코어가 많으면 동시성 최대 16까지 활용.

## 3. 회귀 안전망 (가장 먼저, 절대 생략 금지)

리포에 테스트가 0건이다. **§6 Phase 3-0을 가장 먼저** 한다:
1. 리포트 §5.2 순수 유닛 골든 작성 (`_trim_to_plan_slides`, `_SLIDE_EDITORS` 레지스트리, restructure 멱등성) — 코드 변경 전 베이스라인.
2. 리포트 §5.1 엔드투엔드 골든 작성 (frozen plan 2~3개 + `_date.today()`/`SKILL_DIR` monkeypatch).
3. **이후 모든 코드 변경 step은 완료 직후 골든 테스트로 회귀 확인.** 깨지면 그 step 롤백.

## 4. 실행 순서 (리포트 §6 그대로)

`🔒` = worktree 격리 권장. **순서·격리·선행조건을 반드시 지킨다.**

- **3-0 안전망** → **3-1 무위험 삭제**(soffice, edit_slide15 v1+config, tools.json) → **3-2 config 이전**(slide42) → **3-3 최신화**(bedrock 모델ID, adaptive thinking) → **3-4 AHE 관찰성/자가치유**(qa-ok, manifest, evolve, independent-qa — 상호의존 높음, 순차) → **3-5 needs-care**(각 선행조건 충족 후).

- 같은 파일 인접 함수를 건드리는 step(4·6, 8·9, 14·15)은 **동일 worktree에서 순차** 처리해 merge 충돌 회피.
- 각 step 또는 묶음마다 **단계별 커밋** (이 브랜치). 메시지에 후보 ID 명시.

## 5. 절대 건드리지 말 것 (검증으로 확정)

- **반려(§3):** `dead-redlining-validator`, `slide13-subheading-truncate` — 실행 금지.
- **live(제거 금지):** `analyze-qa`(ahe_loop.py), `pack.py`/`unpack.py`.
- **변경 불요:** `slide25-26`(코드), `sdk-usage-current`.
- **보류:** `dead-docx-validator` — docx 경로 실사용 여부(셸/CI/수동 호출) 확인 전 삭제 금지.
- needs-care의 "틀린 전제" 주의: `common-subtitle-shape-9`의 8M/14pt는 **틀린 값**(실측 cx=3,017,960 / 10.5pt) → 상수화하면 truncation 영구 파손. shape_id 이전만 하고 실측값으로 수정.

## 6. 완료 후

- 단계별 커밋 → `git push origin refactor/2026-05-modernization`.
- 변경 요약 + 골든 테스트 결과(통과/실패)를 보고. 실패 step은 명시.
- (선택) PR 생성: base `main` ← `refactor/2026-05-modernization`.

## 7. 보안 주의

- 이 repo의 git remote URL에 GitHub 토큰이 평문으로 박혀 있을 수 있다. 작업 후 해당 토큰 **폐기(rotate)** 후 credential helper/SSH로 전환 권장.

---

### 윈도우 세션에 붙여넣을 프롬프트 (요약)

> 이 repo의 `PHASE3_HANDOFF.md`와 `REFACTOR_DIAGNOSE_REPORT.md`, `REFACTOR_DIAGNOSE_RAW.json`을 읽어라. 분석/검증은 맥에서 이미 끝났으니 재실행하지 마라. 너는 Phase 3(실제 코드 수정)만 수행한다. PHASE3_HANDOFF.md §3(골든 테스트 먼저)→§4(실행 순서)→§5(금지 항목)를 그대로 따르고, Sonnet 기본·사고형 step만 Opus로 하이브리드 진행해라. 각 step 후 골든 테스트로 회귀 확인하고 단계별 커밋한다. 시작 전 Plan을 보여주고 내 승인을 받아라.
