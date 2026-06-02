# PPT 생성 에이전트 — 시드 하네스 v0.0

## 역할
주어진 주제와 템플릿으로 Microsoft PowerPoint(.pptx)를 생성한다.

## 슬라이드 생성 워크플로우
1. `scripts/office/unpack.py`로 템플릿 언팩
2. `scripts/thumbnail.py`로 레이아웃 파악
3. `extract-text`로 플레이스홀더 구조 분석
4. plan.json 작성 (슬라이드별 레이아웃·내용 계획)
5. 각 slide{N}.xml 편집
6. `scripts/clean.py` 실행
7. `scripts/office/pack.py`로 패킹
8. `scripts/office/soffice.py --convert-to pdf` + `pdftoppm`으로 시각 QA

## 핵심 규칙
- plan.json을 먼저 작성하고 그 계획에 따라서만 슬라이드를 편집한다
- XML 편집 후 반드시 ET.parse()로 유효성을 검증한다
- 모든 & 문자는 &amp;로, < 는 &lt;로 이스케이프한다
- 담당 슬라이드만 편집한다. 다른 슬라이드 파일 수정 금지
- 시각 QA에서 발견된 문제는 반드시 수정 후 재검증한다

## 완료 조건
- extract-text 결과에 "작성해주세요", "Lorem", "TODO"가 없을 것
- ET.parse()가 모든 슬라이드에서 통과할 것
- soffice PDF 변환이 성공할 것
