# PPT 생성 에이전트 — GS Neotek 2026 템플릿 (v1.0)

## GS Neotek 디자인 가이드 (Design Guide p.1~2)
- COLOR: #1419AB(네이비), #3C41E6(블루), #FF4B4B(강조 레드), #DDDDDD(회색)
- FONT: Pretendard SemiBold(제목/목차/장표타이틀), Medium(설명타이틀), Regular(설명상세)
- SIZE: 최소 12pt, 최대 48pt, 짝수 단위(12→14→16→18→20...→48pt)
- LOGO: GS Neotek 로고는 수정 금지 — 슬라이드 하단 고정 위치 유지

## 슬라이드 구성 원칙
- N장 = 표지(1) + 목차(1) + 본문(N-3) + 감사합니다(1)
- closing(slide46): 편집 금지 — "감사합니다." 레이아웃 내장
- section(slide8): 사용자가 명시적으로 요청할 때만 사용
- 슬라이드 순서 = plan index 순서 (PPTX sldIdLst 재정렬 필수)

## XML 편집 핵심 규칙
- Shape는 반드시 ID로 찾는다 (XML 순서 의존 금지)
- `<a:r>`은 `<a:endParaRPr>` 앞에 위치해야 함 → `para.insert(idx, r_new)` 사용
- 한국어 텍스트: `lang="ko-KR" altLang="en-US"` 필수 (en-US → ?? 렌더링)
- rPr은 새로 생성 금지 — `copy.deepcopy(원본rPr)`로 색상/폰트 보존
- presentation.xml 수정: ET.tostring() 금지 → regex string 조작만 허용

## 완료 조건
- placeholder 잔여 없음("작성해주세요", "lorem", "1.4", "이미지/ 영상" 등)
- PowerPoint PDF QA 통과 (LibreOffice 폴백 사용 금지)
- 섹션 구조: 표지/목차·간지/본문/마무리 4개
