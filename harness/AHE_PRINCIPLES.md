# AHE 하네스 엔지니어링 원칙 (2026-06 공식 출처 기반)

> 이 문서는 본 프로젝트의 모든 하네스 변경 판단 기준이다.
> 즉흥 설계 금지 — 아래 원칙과 충돌하면 변경을 보류하고 근거를 제시한다.
> 출처는 2026년 6월 기준 1차 자료(Anthropic 엔지니어링 블로그 + AHE 논문 arXiv:2604.25850).

---

## 0. 출처 (primary sources)
- Anthropic — *Harness design for long-running application development* (anthropic.com/engineering/harness-design-long-running-apps)
- Anthropic — *Scaling Managed Agents: Decoupling the brain from the hands* (anthropic.com/engineering/managed-agents)
- *Agentic Harness Engineering: Observability-Driven Automatic Evolution of Coding-Agent Harnesses* — arXiv:2604.25850
- Official AHE code — github.com/china-qijizhifeng/agentic-harness-engineering

---

## 1. 핵심 원칙 (Anthropic 공식)
1. **constrain / inform / verify / correct / human-in-loop** — 에이전트가 할 수 있는 것을 제약하고, 해야 할 것을 알리고, 작업을 검증하고, 실수를 교정하고, 고위험 결정엔 사람을 둔다.
2. **생성자 ≠ 판정자 (Separation of concerns)** — "생성하는 에이전트와 판정하는 에이전트를 분리한다." 검증 품질은 분리에서 나온다.
3. **명시적·채점 가능 기준** — 주관적 요구를 "concrete, gradable terms"로 변환해야 검증이 신뢰 가능.
4. **작업 분해 + 구조화된 핸드오프** — 장기 작업은 다룰 수 있는 단위로 쪼개고, 세션 간 컨텍스트는 구조화된 산출물로 넘긴다.

## 2. 확증편향 방지 (CONFIRMATION BIAS — 공식 경고)
- Anthropic 명시: Claude는 자기 결과물을 **"사람 눈엔 명백히 평범한데도 자신 있게 칭찬하는 경향"**이 있다.
- 따라서 **"생성자가 자기 작업을 비판하게 만드는 것보다 외부(분리된) 평가 루프가 훨씬 더 다룰 만하다."**
- **프로젝트 규칙**: QA/리뷰는 **생성 컨텍스트를 공유하지 않는 독립 격리 에이전트**가 수행한다. 리뷰어는 plan·생성 근거를 모르고, 결과물(렌더 이미지) + 원 주제/요구만 본다.
- 평가자는 채점 기준표(few-shot 점수 분해)로 **보정(calibrate)**해 점수 드리프트를 막는다.

## 3. stale 가정 제거 (현 모델 수준 — 공식 권장)
- Anthropic 명시: **"하네스의 모든 구성요소는 '모델이 혼자 못하는 것'에 대한 가정을 인코딩한 것이며, 이 가정은 모델이 좋아지면 빠르게 stale해지므로 stress test 해야 한다."**
- 사례: Opus 4.6에서 4.5가 필요로 하던 sprint 분해·context reset을 **"dead weight"로 제거**.
- **프로젝트 규칙 (Opus 4.8 기준)**: 하네스 컴포넌트를 추가/유지하기 전 자문한다 —
  **"이 scaffolding이 현 모델(Opus 4.8)에서 still load-bearing인가? 모델이 이미 할 수 있는 일을 대신 강제하고 있지 않은가?"**
  - 약한 모델용 보상 장치(하드코딩 fix 누적본, 과잉 planning_constraints, rigid fallback)는 stress test 후 비-load-bearing이면 **제거**한다.
  - 단, 결정적 정확성이 필요한 영역(shape ID·좌표·XML 규칙)은 모델이 잘해도 **하네스에 고정**(재현성).

## 4. falsifiable-contract — 모든 하네스 변경의 의무 (AHE 결정 관찰성)
- AHE 명시: **"완전 자동 루프가 없어도 이 패턴을 채택하라 — 모든 하네스 변경은 '어떤 실패를 고칠 것으로 기대하는가, 어떤 동작이 회귀할 수 있는가'를 글로 답해야 한다."**
- **각 편집은 다음 평가로 falsifiable해진다** — 근거 기반 자기정당화를 라운드 간 측정 가능한 계약으로 대체.
- **프로젝트 규칙**: 모든 하네스 변경은 `change_manifest.jsonl`에 1건 기록한다. 필드:
  - `evidence` (어떤 실패/관찰이 근거인가)
  - `root_cause` (추론한 근본 원인)
  - `fix` (무엇을 바꿨나, 파일·라인)
  - `predicted_fixes` (이 변경이 고칠 것)
  - `regression_risk` (회귀할 수 있는 동작)
  - `verification` (다음 실행에서 어떻게 검증됐나 — pending→verified/refuted)
- 이 기록이 곧 **사용자 통보**다 ("업데이트마다 알려줘"의 구현).

## 5. AHE 3기둥 정본 정의 (component / experience / decision)
- **① 컴포넌트 관찰성**: 편집 가능한 하네스 컴포넌트를 **파일**로 노출(고정 위치, git diff·rollback). 본 프로젝트: `harness/*.json`, `SKILL.md`, sub-agent 프롬프트.
- **② 경험 관찰성**: 실행 트레이스를 **계층적 파일**(per-run 리포트 + 전체 overview + raw)로 distill — progressive disclosure로 토큰 절약.
- **③ 결정 관찰성**: 편집 + 예측 manifest(§4), 다음 라운드에서 fix precision/recall로 검증. "evidence ledger".

## 6. 브레인/핸드 분리 (Managed Agents)
- **Brain**(Claude+하네스) / **Hands**(샌드박스·도구) / **Session**(append-only 로그)를 분리. 인터페이스는 "모양에만 opinionated, 구현엔 불가지".
- **프로젝트 규칙**: python 엔진은 **결정적 hands**(XML 편집·패킹·렌더)로 한정. plan·리뷰 같은 **brain 작업은 오케스트레이터 세션/독립 에이전트**가 수행(인라인 LLM 호출은 headless 폴백으로만).

---

## 적용 체크리스트 (모든 하네스 PR/변경 전)
- [ ] 이 변경이 현 모델 기준 load-bearing인가? (§3 stress test)
- [ ] 생성과 판정이 분리돼 있는가? (§2)
- [ ] `change_manifest.jsonl`에 falsifiable-contract를 기록했는가? (§4)
- [ ] 하네스로 표현 가능한가? (불가피한 경우만 Python, 파라미터는 하네스에서)
- [ ] 사용자에게 통보했는가?
