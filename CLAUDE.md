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

| AHE 기둥 | 구현 위치 | 역할 |
|---------|-----------|------|
| ❶ Component Observability | `harness/` 파일들 | 하네스를 파일로 분리, git 추적 |
| ❷ Experience Observability | `ahe_tools/distill_digest.py` | 실행 트레이스 → 구조화된 digest |
| ❸ Decision Observability | `evolution/iteration_N_manifest.json` | 편집+예측 선언 → 다음 라운드 검증 |
