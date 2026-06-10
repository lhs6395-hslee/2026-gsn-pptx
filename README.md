# PPT 생성 스킬

Claude Code에서 주제 한 줄로 Microsoft PowerPoint를 자동 생성합니다.  
AHE(Agentic Harness Engineering) 기법으로 하네스가 실행마다 스스로 개선됩니다.

---

## 설치 (한 번만)

```bash
git clone <this-repo> ppt-skill
cd ppt-skill
bash setup.sh

# 템플릿 배치
cp your_template.pptx ~/.ppt-skill/templates/default.pptx

# (선택) 터미널 alias
echo "alias ppt='~/.ppt-skill/bin/ppt'" >> ~/.zshrc && source ~/.zshrc
```

---

## Claude API 백엔드 설정

Anthropic API 키 없이 **Team Plan**을 사용하는 경우 Vertex AI 또는 Bedrock을 백엔드로 씁니다.  
`~/.claude/settings.json`의 `env` 블록에 값이 있으면 **자동으로 감지**합니다.

### 자동 감지 우선순위

```
ANTHROPIC_VERTEX_PROJECT_ID 존재  →  Vertex AI
CLOUD_ML_REGION만 존재            →  Vertex AI
AWS_REGION만 존재                 →  Bedrock
ANTHROPIC_API_KEY 존재            →  Anthropic 직접
```

> **둘 다 설정된 경우:** `ANTHROPIC_VERTEX_PROJECT_ID`가 있으면 Vertex 우선.  
> 명시적으로 지정하려면 `--backend` 옵션을 사용하세요.

### 백엔드별 필요 환경변수

| 백엔드 | 필수 환경변수 | 비고 |
|--------|-------------|------|
| **Vertex AI** | `ANTHROPIC_VERTEX_PROJECT_ID`, `CLOUD_ML_REGION` | GCP 프로젝트 ID + 리전 |
| **Bedrock** | `AWS_REGION` | AWS IAM 자격증명도 필요 |
| **Anthropic** | `ANTHROPIC_API_KEY` | 개인 키 보유 시 |

`~/.zshrc`에 직접 추가하거나 `~/.claude/settings.json`의 `env` 블록에 넣으면 됩니다.

```jsonc
// ~/.claude/settings.json
{
  "env": {
    "ANTHROPIC_VERTEX_PROJECT_ID": "your-gcp-project-id",
    "CLOUD_ML_REGION": "global",
    "ANTHROPIC_DEFAULT_SONNET_MODEL": "claude-sonnet-4-6"
  }
}
```

---

## 실행 방법

### 1. Claude Code 스킬 (가장 자연스러운 방식)

```
> Kafka 아키텍처 PPT 10장 만들어줘
> React 입문 발표 자료 8장, 대상: 프론트엔드 개발자
> AI 전략 슬라이드 12장 --evolve
```

### 2. Python 직접 실행 (터미널 — 추후 확장)

> 엔진(`ppt_generator.py`, `main.py`)은 `~/.ppt-skill/`에 설치된다.
> 현재는 Claude Code 스킬 경로를 주로 사용하고, 터미널 실행은 추후 확장 예정.

```bash
cd ~/.ppt-skill

# 기본 (백엔드 자동 선택)
python3 main.py --topic "Kafka 아키텍처"

# 슬라이드 수·청중 지정
python3 main.py --topic "Kafka 아키텍처" --slides 10 --audience "백엔드 개발자"

# 백엔드 명시
python3 main.py --topic "Kafka 아키텍처" --backend vertex
python3 main.py --topic "Kafka 아키텍처" --backend bedrock
python3 main.py --topic "Kafka 아키텍처" --backend anthropic

# AHE 진화 루프 활성화
python3 main.py --topic "Kafka 아키텍처" --evolve

# 출력 위치 지정 (기본: ~/Desktop)
python3 main.py --topic "Kafka 아키텍처" --output ~/Documents
```

### 3. 터미널 alias (alias 설정 후)

```bash
ppt "Kafka 아키텍처"
ppt "React 입문" 8 "프론트엔드 개발자"
ppt "AI 전략" 12 "임원진" --evolve
```

---

## --backend 옵션 상세

```
--backend auto       기본값. settings.json → 환경변수 순서로 자동 감지
--backend vertex     Google Vertex AI 강제 (ANTHROPIC_VERTEX_PROJECT_ID 필요)
--backend bedrock    AWS Bedrock 강제 (AWS_REGION + IAM 자격증명 필요)
--backend anthropic  Anthropic API 직접 (ANTHROPIC_API_KEY 필요)
```

환경변수 `PPT_SKILL_BACKEND`로도 같은 효과를 낼 수 있습니다:

```bash
PPT_SKILL_BACKEND=bedrock python3 main.py --topic "..."
```

---

## 실행 흐름

```
주제 입력
  │
  ▼
1. ~/.ppt-skill/templates/default.pptx 언팩
2. 슬라이드 레이아웃 분석 (46개 슬라이드 구조 파악)
3. Claude API → plan.json 생성 (슬라이드별 제목·내용·레이아웃 결정)
   └─ 폴백: Claude 미응답 시 규칙 기반 계획
4. 각 슬라이드 XML 편집 (제목·내용 교체)
   └─ 차트 레이아웃(slide40/41/43)은 content.chart_data를 차트 캐시 + 임베디드 xlsx에 주입
5. clean.py → pack.py → output.pptx 생성
6. PowerPoint(AppleScript) PDF 변환 → pdftoppm 이미지 → qa_report.json
   └─ LibreOffice 미사용 (PowerPoint 정확 렌더링만 신뢰)
7. extract-text로 플레이스홀더 잔여 검증
8. ~/Desktop/<주제>.pptx 복사
  │
  └─ --evolve 시: 트레이스 → digest → manifest → long_term_memory 업데이트
```

---

## 차트 슬라이드 (Excel 연동)

slide40/41/43은 임베디드 차트 레이아웃으로, plan의 `content.chart_data`를 차트 캐시
(`chartN.xml`)와 연동된 임베디드 워크북(`embeddings/*.xlsx`)에 주입해 템플릿 더미값을
실제 수치로 교체한다. ⚠️ 데이터포인트 **개수는 템플릿 고정**(c:pt 추가/삭제 시 PowerPoint
PDF export가 무음 실패) — 아래 개수에 맞춰 제공한다.

| 슬라이드 | 차트 | `chart_data` 스키마 |
|---------|------|--------------------|
| **slide40** | 도넛 3개 | 리스트 3항목, 각 `{title, values:[정확히 3개]}` (범주 라벨 미표시) |
| **slide41** | 막대(다시리즈) | `[{categories:[4개], series:[{name, values:[4개]} … 최대 3]}]` |
| **slide43** | 막대(시계열) | `[{categories:[6개], series:[{name, values:[6개, 음수 가능]} × 2]}]` |

매핑은 `harness/chart_map.json`에 정의된다. slide43(chart5)은 OLE 객체라 임베디드 xlsx가
없어 차트 캐시만 갱신한다.

---

## 파일 구조

```
~/.ppt-skill/                   ← 설치 위치
├── SKILL.md                    ← Claude Code 스킬 선언
├── harness/
│   ├── CLAUDE.md               ← 에이전트 시스템 프롬프트
│   ├── AHE_PRINCIPLES.md       ← 하네스 엔지니어링 원칙 (공식 출처 기반)
│   ├── change_manifest.jsonl   ← 변경별 falsifiable-contract 원장 (결정 관찰성 ❸)
│   ├── long_term_memory.json   ← 누적 경험 (AHE 자동 업데이트)
│   ├── slide_catalog.json      ← verified/unverified/banned 3단계 레이아웃 분류
│   ├── chart_map.json          ← 차트 레이아웃(40/41/43)→차트파일·타입 매핑
│   └── verifier_rules.json     ← 검증 기준
├── scripts/office/
│   ├── unpack.py               ← pptx → 디렉토리
│   └── pack.py                 ← 디렉토리 → pptx
│                               (시각 QA PDF 변환은 PowerPoint AppleScript 사용)
├── templates/
│   └── default.pptx            ← ← 여기에 템플릿 배치
├── runs/                       ← 실행별 작업 디렉토리
│   └── 20260601_150000_Kafka/
│       ├── plan.json           ← Claude가 생성한 슬라이드 계획
│       ├── output.pptx
│       ├── qa_report.json      ← 시각 QA 결과
│       └── qa_images/          ← 슬라이드 이미지 (46장)
├── traces/                     ← AHE 실행 트레이스
├── evolution/                  ← AHE 변경 매니페스트
│
├── ppt_generator.py            ← 핵심 엔진 (분석·계획·편집·패킹·QA)
├── analyze_zones.py            ← 템플릿 존 자동 분류 → layout_zone_map.json
├── main.py                     ← CLI 진입점 (터미널, 추후 확장)
└── ahe_loop.py                 ← AHE 세 기둥 구현
```

> 이전의 별도 `ppt_harness_project/` 폴더는 `ppt-skill`로 통합되었다 (엔진이 스킬과 한 곳에서 동작).

---

## AHE가 하는 일

`--evolve` 옵션을 쓰면:

1. 각 슬라이드 생성 결과를 `traces/`에 기록
2. 실패 패턴 분석 → `evolution/iteration_*_manifest.json` 생성
3. `harness/long_term_memory.json`에 경험 누적 (git으로 변경 추적)
4. 다음 실행 시 같은 실수를 반복하지 않음

---

## 환경 요구사항

| 항목 | 버전 | 설치 |
|------|------|------|
| Python | 3.11+ | — |
| anthropic | 최신 | `pip install "anthropic[vertex]"` |
| python-pptx | — | `pip install python-pptx` |
| Microsoft PowerPoint | macOS | 시각 QA PDF 렌더용 (AppleScript). **LibreOffice 미사용** |
| Poppler | — | `brew install poppler` (pdftoppm) |
| Claude API | Team Plan | Vertex AI 또는 Bedrock 설정 필요 |
