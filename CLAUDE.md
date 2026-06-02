# PPT AHE 하네스 프로젝트

> Claude Code에서 이 파일을 읽으면 이전 작업 컨텍스트를 그대로 이어받을 수 있다.

---

## 프로젝트 목표

Microsoft PowerPoint(.pptx)를 **주제만 입력하면 자동 생성**하는 시스템.
AHE(Agentic Harness Engineering, arXiv:2604.25850) 기법을 적용해
**하네스 컴포넌트가 실행마다 스스로 개선**된다.

핵심 원칙:
- Python 파일은 설치 시 1회만 생성. 주제가 바뀌어도 코드는 그대로.
- 매 실행 시 `runs/<날짜>_<주제>/` 작업 폴더만 새로 생성됨.
- Claude Code 스킬로 동작 — `python main.py`를 직접 실행하지 않음.

---

## 디렉토리 구조

```
ppt-ahe-project/
│
├── CLAUDE.md                   ← 이 파일 (Claude Code 컨텍스트)
│
├── ppt-skill/                  ← Claude Code 스킬
│   ├── SKILL.md                ← 스킬 선언 (트리거 조건, 실행 흐름 정의)
│   ├── setup.sh                ← 한 번만 실행하는 설치 스크립트
│   ├── harness_hooks.py        ← XML 편집 Pre/Post 훅
│   ├── bin/ppt                 ← 터미널 alias용 실행 스크립트
│   ├── harness/                ← AHE 진화 대상 컴포넌트 (7개 파일)
│   │   ├── CLAUDE.md           ← 에이전트 시스템 프롬프트 (v0.3, 레이아웃 카탈로그 포함)
│   │   ├── tools.json          ← 도구 정의
│   │   ├── middleware.py       ← Pre/PostToolUse 규칙
│   │   ├── verifier_rules.json ← 검증 기준 (플레이스홀더, XML, 시각 QA)
│   │   ├── long_term_memory.json ← 누적 경험 (AHE 루프가 자동 업데이트)
│   │   ├── layout_features.json ← 템플릿 슬라이드 구조 메타데이터 (v1.0, 41슬라이드)
│   │   └── thumbnails/         ← 슬라이드별 썸네일 PNG (레이아웃 시각 확인용)
│   ├── scripts/                ← PPTX 조작 스크립트 (Anthropic 스킬)
│   │   └── office/
│   │       ├── unpack.py       ← pptx → 디렉토리
│   │       ├── pack.py         ← 디렉토리 → pptx
│   │       └── soffice.py      ← PDF 변환
│   └── ahe_tools/
│       └── distill_digest.py   ← 트레이스 → digest 압축 (AHE ❷)
│
├── ppt_harness_project/        ← Python 구현체 (Claude Code 없을 때 폴백)
│   ├── main.py                 ← CLI 진입점
│   ├── ppt_generator.py        ← PPT 생성 로직 (분석→계획→편집→검증→패킹)
│   ├── ahe_loop.py             ← AHE 세 기둥 구현
│   └── harness/                ← 동일한 하네스 컴포넌트
│
└── templates/                  ← .pptx 템플릿 보관 위치
    └── (여기에 템플릿 배치)
```

---

## 핵심 기술 결정 사항

### 왜 Claude Code 스킬인가
- `python main.py --topic "..."` 방식은 매번 새 프로세스 실행 → 오버헤드
- 스킬은 Claude 자신이 오케스트레이터 → Python은 도구로만 호출
- 주제가 달라져도 파이썬 파일은 전혀 건드리지 않음

### AHE 세 기둥 (arXiv:2604.25850)
| 기둥 | 구현 위치 | 역할 |
|------|-----------|------|
| ❶ Component Observability | `harness/` 5개 파일 | 하네스를 파일로 분리, git 추적 |
| ❷ Experience Observability | `ahe_tools/distill_digest.py` | 실행 트레이스 → 구조화된 digest |
| ❸ Decision Observability | `evolution/iteration_N_manifest.json` | 편집+예측 선언 → 다음 라운드 검증 |

### 에이전트 분리 수준
- 모델 가중치: 분리 안 됨 (모두 동일한 Claude)
- 컨텍스트 창: 완전 격리 (서브에이전트는 자기 슬라이드만 봄)
- 파일시스템: 공유하되 권한으로 제한 (훅 레이어)
- 통신: 파일 매개 (plan.json, qa_report.json)

---

## 이전 작업에서 만들어진 산출물

### 실제로 생성한 PPT
- Kafka 심층 아키텍처 가이드 (10장, GS Neotek 템플릿 기반)
- 템플릿: `2026_PPT_Template.pptx` (GS Neotek 브랜드, 46장 슬라이드)
- 브랜드 색상: `#1419AB` (네이비), `#3C41E6` (블루), `#FF4B4B` (강조 레드)
- 폰트: Pretendard (SemiBold/Medium/Regular/Light)

### 레이아웃 선택 파이프라인 (v0.3)

레이아웃 선택은 반드시 3단계를 순서대로 실행:
1. `~/.ppt-skill/harness/layout_features.json` → 알고리즘 필터 (신뢰도·컬럼수·아이콘수 매칭)
2. `~/.ppt-skill/harness/thumbnails/slide{NN}.png` → 썸네일 시각 확인 (Read로 이미지 로드)
3. `~/.ppt-skill/harness/long_term_memory.json` → 과거 성공/실패 사례 참조

### 템플릿 슬라이드 카탈로그 (시각 검증 완료, 2026-06-02)

신뢰도: ✅high 🟡medium 🔶medium_low(이미지필요) 🔴low(차트고정)

| 슬라이드 | 역할 | 핵심 구조 | 신뢰도 | 최적 콘텐츠 유형 |
|---------|------|---------|--------|--------------|
| slide6 | 표지 | 날짜+대제목(48pt)+부제목(32pt)+네이비패널 | ✅ | 발표 첫 장 |
| slide7 | 목차 | 번호+항목 7개, 좌측 발표제목 | ✅ | 전체 목차 |
| slide8 | 섹션구분 | 전체 다크 네이비, 좌 섹션제목+우 소목차 | ✅ | 챕터 전환 |
| slide13 | 3열 아이콘카드 | 아이콘3+라운드제목+설명+Insight배너 | 🟡 | 기능/장점 3가지 |
| slide14 | 4열 아이콘카드 | 아이콘4+제목+설명 | 🟡 | 기능/구성요소 4가지 |
| slide15 | 3열 대형아이콘 | 대형아이콘3(중앙)+제목+설명 | 🟡 | 핵심 가치 3가지 |
| slide16 | 4열 아이콘컴팩트 | 아이콘4(소)+제목+설명 | 🟡 | 구성요소 4가지 간결 |
| slide24 | 2블록 텍스트 | 텍스트 2블록, 이미지 없음 | ✅ | 짧은 2가지 비교 |
| slide29 | 연도별 타임라인 | 가로바 2023→2026 | ✅ | 로드맵/연혁 |
| slide30 | 4단계 스텝 | Step1→Step2→Step3→Step4 박스 | ✅ | 프로세스/절차 |
| slide31 | 분기별 타임라인 | Q1→Q2→Q3→Q4 마일스톤 | ✅ | 분기 계획 |
| slide32 | 텍스트+배너 | 상단배너+전체텍스트 | ✅ | 긴 설명/부록 |
| slide34 | 4버블 키워드 | 타원4개+연결화살표 | 🟡 | 연관 개념 맵 |
| slide35 | Before→After 버블 | 소버블군집→대형버블 | 🟡 | 이관/통합 효과 |
| slide36 | As-is/To-be 벤다이어그램 | 좌:As-is 우:To-be+설명 | 🟡 | 현황→목표 비교 |
| slide37 | 막대그래프 비교 | 낮은막대 vs 높은막대+% | 🟡 | 정량적 개선 수치 |
| slide38 | 3행 흐름도 | [keyword]→[Solution]→[Service] × 3행 | ✅ | 아키텍처/파이프라인/매핑 |
| slide39 | 3행 상세 흐름도 | slide38+세부항목 | ✅ | 복잡한 흐름+세부항목 |
| slide40 | 도넛차트3개 | 도넛3+인사이트배너 | 🔴 | KPI 3가지 |
| slide41 | 라인차트+콜아웃 | 꺾은선+레이블3 | 🔴 | 트렌드 추이 |
| slide42 | 대형 도넛3개 | 도넛3(큰)+제목+설명 | 🔴 | 핵심 KPI 임팩트 |
| slide43 | 라인차트+2인사이트 | 그래프+설명2블록 | 🔴 | 분석결과+인사이트 |
| slide44 | Q&A | 풀스크린블루+"Q&A" | ✅ | 질의응답 |
| slide46 | 감사합니다 | 블루그라디언트+"감사합니다." | ✅ | 마지막 슬라이드(편집금지) |

### 콘텐츠 유형별 최우선 추천
```
아키텍처/파이프라인 매핑 → slide38 (가장 높은 신뢰도)
기능/특징 3가지          → slide13 또는 slide15
기능/특징 4가지          → slide14 또는 slide16
4단계 프로세스           → slide30
로드맵 연도별            → slide29
현황→목표 비교           → slide36 (개념) / slide37 (수치)
이관/전환 효과           → slide35 (Before→After 버블)
```

### 더 좋은 방법 (개선 방향)

현재 카탈로그 방식의 한계와 개선 방향:

| 현재 방법 | 한계 | 개선 방향 |
|----------|------|---------|
| 텍스트 카탈로그 | 애매한 설명, 오해 가능 | 썸네일 시각 확인 (현재 적용) |
| 수동 layout_features.json | 템플릿 교체 시 구식화 | `analyze_template.py` 자동 재생성 |
| 직관적 선택 | 경험 없으면 오선택 | long_term_memory로 성공률 누적 |
| 생성 후 QA | 오류 발견이 늦음 | 선택 시점에 capacity 검증 추가 |

### 템플릿 작업 시 알려진 이슈 & 해결법
1. **목차(slide7) 번호 정렬**: 번호 박스와 내용 박스가 별도 shape. `anchor="ctr"` + `spcPct val="150000"` 맞춰야 정렬됨
2. **섹션(slide8) 서브아이템 정렬**: `anchorCtr="0"` 추가 필요
3. **헤더 중복**: `>PPT <` 교체 후 다음 run의 `>Kafka<` 중복 발생 → regex로 제거
4. **XML 이스케이프**: `&` → `&amp;`, `<` → `&lt;` 필수. 미처리 시 pack 실패
5. **lorem ipsum 잔여**: 여러 `<a:r>` run에 분산됨. 단순 replace 아닌 regex sub 필요

---

## TODO — 다음 작업 항목

### 즉시 해야 할 것 (완료)
- [x] `ppt_generator.py`의 `run_ppt_generation()` 함수 완성 (2026-06-01)
  - `analyze_template` → `generate_plan` → `edit_slide` 루프 구현
- [x] `setup.sh` + end-to-end 테스트 통과 (Kafka 아키텍처 5장, 2026-06-01)
- [x] AHE 루프 첫 실행 (`--evolve`) 검증 (React 입문 5장 + evolve 성공, 2026-06-01)

### 이후 작업 (완료)
- [x] `harness/long_term_memory.json` v0.1 — GS Neotek 패턴 추가 (slide7 anchor=ctr, slide8 anchorCtr=0, XML 이스케이프, lorem 잔여 fix, 2026-06-01)
- [x] `harness/verifier_rules.json` v0.1 — toc_alignment, section_anchor, xml_escape, visual_overflow known_patterns 추가 (2026-06-01)
- [x] LibreOffice 설치 + 시각 QA(PDF→이미지) 파이프라인 완성 (2026-06-01)
- [x] Claude API (Vertex AI 포함) 콘텐츠 생성 경로 구현 및 테스트 통과 (2026-06-01)

### 남은 개선 과제

#### 레이아웃 선택 정확도 개선
- [ ] `analyze_template.py` 구현 — 템플릿 XML 파싱해 layout_features.json 자동 재생성 (setup.sh에 포함)
- [ ] 콘텐츠→레이아웃 매칭 함수 — item_count, has_icon, flow_type 등을 입력받아 후보 슬라이드 점수화하는 Python 함수
- [ ] long_term_memory.json에 레이아웃 선택 성공/실패 자동 기록 (QA 점수 연계)

#### 콘텐츠 생성 품질 개선
- [ ] AHE Evolve Agent: manifest 기반으로 harness/*.json 실제 편집하는 로직 구현
- [ ] `edit_slide` 고도화: Claude 계획의 bullets/body 콘텐츠를 XML에 실제 반영 (현재 title만 교체)
- [ ] 시각 QA 이미지를 Claude Vision으로 분석해 오버플로우 자동 감지

#### 인프라 개선
- [ ] slide6 표지 커버 수정: 자간 1.0, 동적 여백 조정, 부제목 폰트 깨짐 해결 (진행 중)

---

## 환경 & 의존성

```bash
# Python
pip install anthropic defusedxml Pillow python-pptx

# 시스템
brew install --cask libreoffice  # soffice (PDF 변환)
brew install poppler              # pdftoppm (이미지 변환)

# 환경변수
export ANTHROPIC_API_KEY="sk-ant-..."
```

---

## 사용법 (설치 후)

```bash
# Claude Code에서 자연어로
> Kafka 아키텍처 PPT 10장 만들어줘
> React 입문 발표 자료 8장, 대상: 프론트엔드 개발자

# 터미널에서
ppt "Kafka 아키텍처"
ppt "React 입문" 8 "프론트엔드 개발자"
ppt "AI 전략" --evolve   # AHE 진화 루프 활성화
```

---

## 참고 자료
- AHE 논문: https://arxiv.org/abs/2604.25850
- AHE GitHub: https://github.com/china-qijizhifeng/agentic-harness-engineering
- Anthropic Multi-agent patterns: https://claude.com/blog/multi-agent-coordination-patterns
