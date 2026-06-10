<!-- 자동 생성: 리팩토링 진단 워크플로 (분석/검증) — 맥 세션 -->
<!-- 생성 환경: macOS (LEEs-MacBook-Pro), Phase1=Opus, Phase2=Sonnet 적대적 검증 -->
<!-- 다음 단계(Phase 3 실행)는 윈도우 네이티브 세션에서 이 리포트 기반으로 진행 -->

# 진단 리포트 (analysis + verification 완료, 코드 미수정)

> 이 문서는 리팩토링 Phase 1(분석)+Phase 2(검증) 산출물입니다.
> Phase 3(실제 코드 수정)는 이 문서의 confirmed/needs-care/rejected 분류를 근거로 진행합니다.
> 기계 참조용 원시 데이터: REFACTOR_DIAGNOSE_RAW.json

---

# 2026-gsn-pptx 리팩토링 진단 — 단일 실행 리포트

6축(dead-code · duplication · harness-migration · ahe-selfheal · modernization · regression-risk) 분석 + 적대적 검증 결과를 verdict별로 분류했다. 총 후보 32건 (전제 정정 4건 포함).

---

## 1. 확정 변경 (confirmed) — 바로 실행 가능

| ID | 위치 | 변경 | 위험 | 신뢰도 |
|----|------|------|------|--------|
| **[A] dead-soffice** | `scripts/office/soffice.py` (전체) | 파일 삭제. import 0건 + SKILL.md:206 soffice 금지 정책 | 낮음 — 런타임 무영향 | high |
| **[A] dead-edit-slide15-v1** | `ppt_generator.py:1563` | 구버전 `edit_slide15()` 본문 삭제 (`_edit_slide15_v2`로 대체됨). 삭제 전 `_load_slide15_config()` 타 참조 grep 확인 | 낮음 — 디스패치는 v2 사용 | high |
| **[A] orphan-tools-json** | `harness/tools.json` | 삭제 + README.md:181 트리 항목 갱신. v0.0 stub, 로더 0건, 초기 커밋 후 무수정 | 낮음 — 로더·소비자 없음 | high |
| **[A] premise: slide15-config** | `harness/slide15_config.json` | **삭제 후보로 재분류.** 적대적 검증이 전제 정정을 *기각* — `_edit_slide15_v2`는 `cfg`/`_load_slide15_config()`를 호출하지 않고 `_load_slide_shape_ids()`만 사용. config + `_load_slide15_config()` + `edit_slide15()` 모두 dead | 낮음 (단 `_SLIDE_EDITORS[slide15.xml]=_edit_slide15_v2` 재확인) | high |
| **[A] premise: analyze-qa** | `ahe_loop.py:81,643` | **제거 금지 — live.** 완전 구현 + `run_evolve_loop`에서 호출 | 제거 시 --evolve Vision QA 소실(HIGH) | high |
| **[A] premise: pack-unpack** | `scripts/office/pack.py·unpack.py` | **제거 금지 — live.** `ppt_generator.py:2217,6601`에서 호출 | 제거 시 파이프라인 붕괴(HIGH) | high |
| **[C] slide42-item-descs-ids** | `ppt_generator.py:5171,5175` | `[\"43\",\"47\",\"51\"]`·`(2_435_125,12,3)`을 `slide_shape_ids.json` slide42에 `item_descs_ids`/`item_desc`로 이전. slide37 패턴. cx=2435125·max_lines=3 실측 정확 확인 | 낮음 — fallback 동일값 무회귀 | high |
| **[C] slide25-26-config-absent** | `ppt_generator.py:5063,5073` | **코드 변경 없음.** zonemap 전담 확인. 단 wrapper 함수는 필수(descriptions→overview/main_desc 분리, 중복 방지). JSON에 `_comment` 문서화만 권장 | 없음 | high |
| **[D] qa-ok-none-manual-close** | `ppt_generator.py:6526,6551` / `SKILL.md:394-399` | 실측: runs 25개 중 24개(96%)가 qa_ok=None → success_rate=1.0이 1샘플 기반 무의미. 자동 닫힘 경로 추가 또는 SKILL.md 11.5를 필수 절차로 강제 | qa_ok 오염, hooks 크로스플랫폼 충돌 | high |
| **[D] evolve-manual-trigger-gap** | `main.py:93-102` / `SKILL.md:392-415` | skill 경로(inline_vision_qa=False)는 `_vision_critical_total=0` 고정 + main.py 미경유 → run_evolve_loop 구조적 도달불가. 독립 QA fail→evolve 트리거 경로 추가 | 자동 진화 회귀, human-in-loop §1 충돌 | high |
| **[D] independent-qa-not-implemented** | `ahe_loop.py:81-107` / `AHE_PRINCIPLES.md:23-27` | ahe_loop는 동일 프로세스 자기-QA(§2 위반). 독립 에이전트는 SKILL.md prose뿐. verdict 스키마 코드화 + headless 폴백 한정 보장 | orchestrator 의존, 비대화형 불가 | high |
| **[D] manifest-ledger-split** | `ahe_loop.py:496-529,674-687` / `change_manifest.jsonl:9` | `.jsonl` write 코드 0건 — verification:'pending' 갱신 경로 없음. iteration_*_manifest와 스키마/소비자 완전 분리. 단방향 브리지로 일원화 | 스키마 통합 시 마이그레이션 | high |
| **[E] stale-bedrock-sonnet-4-5-id** | `ahe_loop.py:218,314` | Bedrock 분기 `claude-sonnet-4-5-20250929-v1:0` 하드코딩(환경변수 무시) → `us.anthropic.claude-sonnet-4-6`로 통일. ppt_generator.py:413-415 패턴 적용 | 낮음 — 4.5 active이나 deprecation+버전 불일치 | high |
| **[E] no-adaptive-thinking-effort** | `ppt_generator.py:438,5791,6015` / `ahe_loop.py:226,322` | Anthropic/Vertex 분기에 `thinking={"type":"adaptive"}`+`output_config={"effort":...}` 추가. Bedrock invoke_model 3곳은 별도 JSON 스키마라 분기 처리 필요 | medium — 400 미유발(선택적), 토큰비용↑ | confirmed* |
| **[E] sdk-usage-current** | `ppt_generator.py:143,431-444` / `ahe_loop.py:187` | **변경 불필요 — 현행 유지.** SDK 생성자·`response.content[0].text`·PEP 604/585 타입 모두 현행. bare except 0건, 타 provider 마커 0건 | 없음 | high |
| **[F] reflection-editor-registry** | `ppt_generator.py:5181-5189,1811-1823` | 레지스트리 완전성 테스트 추가. **추가 발견: slide40/41/43이 `_ALLOWED_CONTENT_SLIDES`에 있으나 레지스트리 부재 → 일반 에디터로 silent fallback.** exact key set assert 필요 | 무진단 fallback | high |
| **[F] harness-autoupdate-side-effect** | `ppt_generator.py:6506-6548` (호출 7027) | `SKILL_DIR/harness/long_term_memory.json`(git-tracked) 무조건 mutate. SKILL_DIR 모듈상수 monkeypatch + 파일 내용 assert(try/except가 실패 삼킴). 격리 테스트 | 테스트가 committed 파일 오염 | high |
| **[F] cleanup-move-output-path** | `ppt_generator.py:7029-7039` | 반환 `tuple[Path,int]`인데 시그니처 `-> Path` stale. **추가 발견: cleanup=True가 rmtree한 work_dir를 main.py:97의 run_evolve_loop가 읽음 → phantom 데이터.** 계약 테스트 + evolve guard | parent.parent 재배치, rmtree 실패 은닉 | high |
| **[F] plan-mutation-chain-units** | `ppt_generator.py:6679-6856` | 10개 순서민감 mutation 블록 단위-골든. `_CHAPTER_MAP_CACHE`/`_SLIDE_CATALOG_CACHE` 글로벌 fixture 리셋, date monkeypatch, 6815 src미존재 시에도 template_file 갱신 동작 핀 | medium — 글로벌 누수, off-by-one | high |
| **[F] pack-restructure-double-call** | `ppt_generator.py:6908,6927,6972` | restructure_sections 외부 명시 호출 3곳 제거(2230 내부 호출이 authoritative). 멱등성(2396 early-exit) 검증됨 | medium — 멱등 확인 후 dedup | high |

\* E축 no-adaptive: 검증 verdict는 confirmed이나 우선순위 medium (기능 노후 아닌 미적용).

---

## 2. 주의 필요 (needs-care) — 조건부 실행

| ID | 위치 | 핵심 조건 / 선행 작업 |
|----|------|---------------------|
| **[A] dead-docx-validator** | `validators/docx.py:16` | "도달 불가"는 과장. `pack.py`/`validate.py`는 **CLI 독립 실행 도구** — `.docx` 인자 직접 호출 가능. 제거 전: (1) 셸/CI/수동 `.docx` 호출 여부 확인, (2) 타 docx 파이프라인 import 확인. 제거 시 `__init__.py __all__` 계약 파괴 → docx.py+redlining.py+simplify_redlines.py+merge_runs.py 일괄 정리해야 일관 |
| **[B] dedupe-nested-set-helpers** | `ppt_generator.py:3134,...,4655 vs 4468` | slide8(L3134)·slide24(L4655)는 **element-arg API**(sid 아님) → 치환 불가. `_slide_set_helper`의 `_auto_resize_textbox`는 사이드바 전용 `_SIDEBAR_LINE_HEIGHT_EMU` 캘리브레이션 → body-zone shape에 오류 cy. `_apply_common_zones._set_shape`(force_semibold+multi-para)·`_apply_slots._set`(spAutoFit 연동)은 동작 상이. **slim `_set_text_run`(cy확장 없음)을 새로 만들어 sid-수용 6곳만 교체**, 위 4곳 제외 |
| **[B] edit-flow-slide** | `ppt_generator.py:4570,4974` | 폴백 체인 3곳 의도적 분기 보존 필수: keywords(slide38=bullets/slide39=items), solutions(items/descriptions), services(slide38=details 포함/slide39=details 제외·detail컬럼 선점). detail 다단락은 paragraph-level rPr deepcopy+잉여단락제거+`_auto_resize_textbox` 미호출 정확 구현. `flow_columns` 스키마 신규. **대안: 공통 3루프만 `_apply_flow_columns` 추출, 폴백+detail은 각 함수 유지(부분 리팩토링이 더 안전)** |
| **[B] edit-timeline-slide** | `ppt_generator.py:3542,3925,3976` | slide31/33만 1차 통합(slide29 제외 타당). (1) auto-registration(`^_edit_(slide\d+)$`) 우회 — thin wrapper `_edit_slide31/33` 유지 필수, (2) `content_top_y` 초기화 순서(slide31=1099874/slide33=1850000)+dynamic_layout 전처리 tmpl_key 분기, (3) 기본값(q_label_ids·col_x_list) 슬라이드별 분리 dict 또는 JSON 값 보장 검증 |
| **[B] fold-slide13-15-into-slots** | `ppt_generator.py:4212,4263,5008` | slide13: 폴백분기 제거 + `_apply_slots`에 `slide_plan=` 추가 후 위임. **slide15 단순 흡수 시 sub_heading(ID=41, zone map 경유) 누락 + Vision Fix 잔여단락 제거 미동작 회귀.** 선행: `slide_shape_ids.json[slide15][slots]`에 `type:sub_heading,ids:["41"]` 추가 또는 `_edit_slots_slide`에 zone map fallback+extra-para-purge 옵션 |
| **[B] extract-content-body-merge** | `ppt_generator.py:4313,...,5134` | 단일 헬퍼 금지. **2개로 분리**: `_merge_body`(setdefault 루프 — slide13/15/_edit_slots_slide 3곳만), `_extract_field(content,body,keys,coerce)`(keys 항상 명시 전달, 추론 금지). slide38/39 services는 키 리스트 다름(`["services","details"]` vs `["services","details2"]`). slide32 body=string 가드 |
| **[C] common-subtitle-shape-9** | `ppt_generator.py:3465` | shape_id 이전(`"9"`→`z.get("subtitle")`)만 안전(35개 body 슬라이드 전부 subtitle:"9" 확인). **단 width_emu=8_000_000(실측 3_017_960의 2.65배)·font_pt=14(실측 10.5pt)는 틀린 값 → config 상수화 시 truncation 영구 파손.** cx=3_017_960/font_pt=10 수정 또는 layout XML 런타임 로딩 필요 |
| **[D] verify-predictions-weak-metric** | `ahe_loop.py:511-522` | 단일 플래그 약점은 실재. **단 제안이 잘못된 데이터 구조 지목** — `predicted_fixes`/`regression_risk`는 `change_manifest.jsonl` 전용 필드, verify_predictions가 읽는 `iteration_*_manifest.json` predictions는 `{change,expected,metric}` 3키뿐. 실제 범위: (1) metric 타입 강제 프롬프트 개선, (2) metric 타입별 분기 |
| **[D] evolve-digest-ignores-qa-none** | `ahe_loop.py:239,262-281` | **success 인플레이션 주장은 틀림** — success(262)는 qa_ok 미포함(`output_exists AND issues==0 AND critical_vision==0`). 실제 문제는 naming 이중성: `qa_ok=image_count>0`이 `flags.qa_done`으로 재노출 → Evolve Agent 오해. rename만 필요: `qa_images_exist`/`headless_qa_images_exist` |
| **[F] golden-deterministic-tail** | `ppt_generator.py:6572,6658` | plan_override 결정성 대부분 맞으나: (1) `_date.today()`(6779) plan_override 무관하게 무조건 실행 → 표지 날짜 매일 변동, monkeypatch 필요, (2) `enforce_plan_constraints`가 banned/MAX_REPEAT=3 초과 template_file in-place 치환 → fixture 주의, (3) SKILL_DIR redirect 필요. 반환 tuple 언팩 |
| **[F] slide-order-rewrite-regression** | `ppt_generator.py:2127-2197` | **호출처 3곳**(6895,6970,**+rebuild_slides.py:415** — 원분석 누락). 정규식 self-closing/Target 가정·멱등성 실측 안전 확인. raw byte 아닌 parsed rId 순서로 assert. 커밋 5872108은 project-steer repo의 spTree 이슈로 본 건과 무관 |
| **[F] vision-dir-detection-fragility** | `ppt_generator.py:6936-6996,6988/7018` | `if "fix_instructions" in dir()` 로컬스코프 의존 — helper 추출 시 NameError/항상False. **단 "inline_vision_qa=False가 main.py 기본" 주장 부정확**(main.py는 미전달→기본 True). 양 경로 계약 테스트 + `fix_instructions=[]` 사전 초기화 또는 명시 플래그 |

---

## 3. 반려 (rejected) — 실행하지 말 것

| ID | 위치 | 반려 이유 |
|----|------|----------|
| **[A] dead-redlining-validator** | `validators/redlining.py:11` | `pack.py`/`validate.py`는 멀티포맷 독립 CLI. `python validate.py dir --original x.docx`로 `.docx` case 실행 시 RedliningValidator 인스턴스화 → 도달 가능. 단독 삭제 시 pack.py:89·validate.py:87 **즉시 ImportError**. dead code 판정 기각. (docx 경로 전체를 먼저 제거하는 선행 결정 없이는 불가) |
| **[C] slide13-subheading-truncate** | `ppt_generator.py:4241` | `(9_000_000,16,2)`는 **dead loop 내부.** `z.get("body",{}).get("sub_heading")`가 `layout_zone_map.json[slide13.xml].body`에 sub_heading 키 부재 → `[]` → for 본문 미실행. slide_shape_ids slots에도 sub_heading 없음. 파라미터 이전해도 런타임 동작 불변 → 무의미. (올바른 순서: zonemap에 키 추가가 선행) |

---

## 4. 2026-05 최신화 갭 요약 (E축)

`[공식]` claude-api SKILL.md model catalog (cached 2026-05-26)

| 항목 | 현황 | 갭 | 조치 |
|------|------|----|------|
| 기본 모델 ID | `ANTHROPIC_DEFAULT_SONNET_MODEL` → `claude-sonnet-4-6` 폴백 | 없음 — 현행 유효 | 유지 |
| **Bedrock 모델 ID** | `ahe_loop.py:218,314`에 `claude-sonnet-4-5-20250929-v1:0` 하드코딩, 환경변수 무시 | **불일치** — 메인 경로(ppt_generator)는 4-6 사용. 4.5는 active이나 legacy·deprecation 대상 | `us.anthropic.claude-sonnet-4-6` 통일 (confirmed) |
| adaptive thinking / effort | 6개 호출 전부 미설정 | **미적용** — Sonnet 4.6 지원하나 400 미유발 | Anthropic/Vertex 분기에 추가, Bedrock 3곳 별도 처리 (medium) |
| 제거된 param (budget_tokens/temperature/top_p) | 코드 어디에도 없음 | 없음 — 4.7/4.8 호환 정상 | 조치 불요 |
| SDK 사용법 | `Anthropic`/`AnthropicVertex`/boto3·`content[0].text`·PEP 604/585 | 없음 — 전부 현행 | 유지 |
| python-pptx/lxml/구버전 관용구 | bare except 0건, 타 provider 마커 0건 | 없음(작업 범위 내) | 유지 |

> SKILL.md "ALWAYS use claude-opus-4-8" 권고 대비 sonnet 계열 의도 선택은 비용 효율상 사용자 결정 영역 — 갭 아님.

---

## 5. 회귀 안전망 제안 (골든 테스트 입출력)

리포에 테스트 0건 (`test_*.py`/`conftest.py`/`pytest.ini` 부재). 결정성 경계는 LLM plan 호출 1곳, 그 이후 전부 결정적.

### 5.1 엔드투엔드 골든 (deterministic tail)
**입력:** `run_ppt_generation(plan_override=<frozen plan.json>, cleanup_work_dir=False, inline_vision_qa=False, template_path=template/'2026_PPT Template.pptx')`
- 필수 전처리: `_date.today()` monkeypatch (표지 날짜 고정), `SKILL_DIR` redirect (long_term_memory 오염 방지), frozen plan은 banned/MAX_REPEAT 미유발 (enforce_plan_constraints 치환 회피)
- 반환 언팩: `output, vision_issues = ...` (tuple, `-> Path` 시그니처 stale)

**출력 핀 (golden fixture):**
1. `verify_content(output) == []` (placeholder 잔여 없음)
2. `[v for v in execute_verifier_rules(output, work_dir) if v['severity']=='CRITICAL'] == []`
3. `check_font_compliance == []`
4. slide별 `extract_slide_text` (byte-stable)
5. `presentation.xml` `sldIdLst` r:id 순서 (parsed, raw 아님)
- 2-3개 frozen plan으로 레이아웃 패밀리 커버: text(slide32) / chart(slide40·41) / timeline(slide29) / comparison(slide35)

### 5.2 순수 유닛 골든 (LLM·PowerPoint 불요)

| 대상 | 입력 | 출력 핀 |
|------|------|---------|
| `_trim_to_plan_slides` (2127) | 스크램블 템플릿 순서 + 다른 plan index 순서 | sldIdLst r:id가 plan 순서 정확 일치 + plan-absent 제거 + **멱등(2회=1회)**. ※rebuild_slides.py:415 경로도 |
| `_SLIDE_EDITORS` 레지스트리 | `import ppt_generator` | `set(keys()) == 하드코딩 expected` (slide8/21/22/28 포함, slide40/41/43 부재 명시). 수동 override 2건(slide22→_edit_slide24, slide15→_edit_slide15_v2) `is` 검증 |
| `_record_run_experience` (6506) | temp memory + monkeypatch SKILL_DIR | runs[] append, total_runs+1, 50-cap, `_success_rate` 산식(non-bool 제외), qa_ok=None 보존. **파일 내용 assert(try/except 실패 은닉)** |
| cleanup 경로 (7029) | cleanup=True/False 양쪽 | (a) 양 분기 2-tuple, (b) True→work_dir 제거+parent.parent/{60자 sanitize}.pptx (space·/→_), (c) False→work_dir/output.pptx |
| plan-mutation 10블록 (6679) | crafted dict | closing 주입/n_slides-3 trim+reindex, dup template→_cN copy, TOC page_nums 균등분배 산식, validate_plan (ok,warnings), enforce_plan_constraints banned 치환. ※`_CHAPTER_MAP_CACHE`/`_SLIDE_CATALOG_CACHE` fixture 리셋 |
| restructure 멱등성 (2369) | pack 1회 vs pack+restructure | 4그룹(표지/목차·간지/본문/마무리) sectionLst 동일 |
| vision 2-mode 계약 (6936) | inline=False / inline=True(mock visual_qa·_run_vision_fix_agent) | 양 경로 vision_issues==0, fix_instructions 미바인딩 시 무크래시 |

---

## 6. Phase 3 실행 권장 순서 (의존성 고려)

`🔒` = worktree 격리 권장 (병렬 충돌·side-effect 위험)

### Phase 3-0 — 안전망 선행 (격리 불요, 가장 먼저)
1. **5.2 순수 유닛 골든 작성** — `_trim_to_plan_slides`, `_SLIDE_EDITORS` 레지스트리, restructure 멱등성. 코드 변경 전 회귀 베이스라인 확보. (의존성 없음, 병렬 가능)
2. **5.1 엔드투엔드 골든 작성** — frozen plan 2-3개 + monkeypatch 전처리. dead-code/리팩토링 검증 기준선.

### Phase 3-1 — 무위험 삭제 (confirmed dead-code, 병렬)
3. 🔒 `dead-soffice` — soffice.py 삭제 (독립)
4. 🔒 `dead-edit-slide15-v1` + `premise:slide15-config` — `edit_slide15()` + `_load_slide15_config()` + `slide15_config.json` 동반 삭제 (`_SLIDE_EDITORS[slide15.xml]=_edit_slide15_v2` 재확인 후). **두 건 동일 파일(ppt_generator.py) 연관 → 같은 worktree에서**
5. `orphan-tools-json` — tools.json + README:181 (독립, 격리 불요)

### Phase 3-2 — 저위험 config 이전 (confirmed, ppt_generator.py 단일 → 순차 또는 같은 worktree)
6. 🔒 `slide42-item-descs-ids` — slide_shape_ids.json + 함수 수정
7. `slide25-26-config-absent` — JSON 주석만 (코드 무변경)

### Phase 3-3 — 최신화 (confirmed, ahe_loop.py/ppt_generator.py)
8. 🔒 `stale-bedrock-sonnet-4-5-id` — ahe_loop.py:218,314 모델 ID 통일
9. 🔒 `no-adaptive-thinking-effort` — 5개 호출 thinking/effort 추가 (Bedrock 3곳 분기). 8번과 같은 ahe_loop.py 건드림 → **8 완료 후 또는 동일 worktree**

### Phase 3-4 — AHE 관찰성/자가치유 (confirmed, 상호 의존 높음 → 격리 필수)
10. 🔒 `qa-ok-none-manual-close` — update_last_run_qa 자동화/강제 (run_id 매칭으로 off-by-one 회피)
11. 🔒 `manifest-ledger-split` — change_manifest.jsonl 단방향 브리지
12. 🔒 `evolve-manual-trigger-gap` — qa_ok=False→evolve 트리거 (human gate 보존). **10·11 결과에 의존 → 10·11 완료 후**
13. 🔒 `independent-qa-not-implemented` — verdict 스키마 코드화 (인터페이스만 우선, spawn은 optional). **12와 연관 → 순차**

### Phase 3-5 — needs-care (조건 충족 후, 격리 필수, 골든 테스트로 회귀 검증)
14. 🔒 `dedupe-nested-set-helpers` — slim `_set_text_run` 신규 + 6곳만 교체 (slide8/24/_apply_common_zones/_apply_slots 제외)
15. 🔒 `extract-content-body-merge` — `_merge_body`+`_extract_field` 분리. **14와 동일 헬퍼 영역 → 14 후 순차**
16. 🔒 `common-subtitle-shape-9` — shape_id 이전만(실측값 수정 동반), 8M/14pt 상수화 보류
17. 🔒 `edit-timeline-slide` — slide31/33 통합(thin wrapper 유지, content_top_y 순서)
18. 🔒 `edit-flow-slide` — 부분 리팩토링(`_apply_flow_columns` 추출, 폴백·detail 유지) 권장
19. 🔒 `fold-slide13-15-into-slots` — slide15 sub_heading 슬롯 JSON 이전 선행 후 흡수
20. `verify-predictions-weak-metric` — metric 타입 강제 + 분기 (범위 재조정)
21. `evolve-digest-ignores-qa-none` — rename만 (qa_images_exist)
22. `pack-restructure-double-call` — 외부 호출 3곳 제거 (멱등성 골든으로 검증 후)
23. `cleanup-move-output-path` + `vision-dir-detection-fragility` — 시그니처 `-> tuple[Path,int]` 수정, evolve work_dir guard, `fix_instructions=[]` 명시 초기화

### 제외 (실행 금지)
- `dead-redlining-validator`, `slide13-subheading-truncate` — **반려** (§3)
- `dead-docx-validator` — docx 경로 사용 여부 확인 전 보류 (CLI 호출 가능성)
- `analyze-qa`/`pack-unpack` — **live, 제거 금지**
- `slide25-26`(코드)·`sdk-usage-current` — 변경 불요

> **격리 원칙:** Phase 3-4(AHE)와 3-5(리팩토링)는 `ppt_generator.py`·`ahe_loop.py`를 동시 다발 수정하므로 worktree 분리 필수. 같은 파일의 인접 함수를 건드리는 step(4·6, 14·15)은 동일 worktree에서 순차 처리해 merge 충돌 회피. 모든 3-5 step은 완료 직후 Phase 3-0 골든 테스트로 회귀 확인.
