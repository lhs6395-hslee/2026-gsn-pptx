# 임원·고객사 보고용 PowerPoint 디자인 가이드

> **기준**: Microsoft PowerPoint · 한국 IT업계 · 임원진/고객사 보고 · 2026년 6월
> **작성 방식**: 모든 규칙은 외부 공식 출처를 직접 열람·확인 후 작성. 출처 신뢰도를 `[공식]`(1차 출처 직접 확인) / `[미확인]`(2차 출처·스니펫 기반)으로 구분.
> **이 프로젝트와의 관계**: PPT AHE 하네스의 디자인 판단 기준 문서. 하네스 규칙(common_formatting.json 등) 변경 시 이 문서를 근거로 판단한다.

---

## 1. 문서 골격 — 한국 대기업 보고서의 사실상 표준

삼성전자 2025년 1분기 경영설명회(15장 전수 확인), 카카오 3Q25 실적발표(16장 중 8장 확인) 실물 PDF에서 직접 추출한 구조. `[공식]`

### 1.1 표준 슬라이드 순서

| 순서 | 슬라이드 | 핵심 규칙 |
|------|---------|----------|
| 1 | **표지** | 회사명/제목 대형 타이틀 + 행사명 부제 + 로고. 어두운 배경(사진 또는 브랜드 단색) |
| 2 | **유의사항(면책)** | 고객사·외부 공개 자료라면 표지 직후. 예측정보 면책, 기준(K-IFRS 등) 고지 |
| 3 | **목차** | 3~4개 대항목을 세로 구분선으로 가로 배열 (삼성 패턴) |
| 4 | **요약 1장 (Executive Summary)** | 핵심 KPI 2~3개를 **대형 숫자 + 단일 강조색**으로 두괄식 제시. 임원은 이 한 장으로 의사결정한다는 전제로 작성 |
| 5~N | **본문 (동일 템플릿 반복)** | "좌측 차트 + 우측 키워드 불릿" 2단 구조. 헤드라인은 결론 수치로 시작 |
| N+1 | **별첨(Appendix)** | 상세 표·깨알 데이터는 본문에서 격리해 별첨으로. 질문 대응용 |
| 끝 | **클로징** | "감사합니다" 단독 장 |

근거:
- 삼성전자 1Q25 경영설명회 PDF — https://images.samsung.com/kdp/ir/events/2025/2025_1Q_conference_kor.pdf `[공식]`
- 카카오 3Q25 실적발표 PDF — https://t1.kakaocdn.net/kakaocorp/admin/ir/results-announcement/5894.pdf `[공식]`

### 1.2 본문 장 공통 패턴 (실물에서 확인된 디테일)

- **섹션 라벨**: 좌상단 캡슐형 라벨("플랫폼 부문 | 톡비즈" — 카카오 패턴) `[공식]`
- **헤드라인 = 결론**: "+9% YoY / +3% QoQ"처럼 수치 결론을 헤드라인에 박는다 `[공식]`
- **시간축 구성**: "현황 → 차분기 전망 → 하반기 전망" 3단 캡슐 헤더 (삼성 패턴) `[공식]`
- **불릿 문체**: `키워드: 설명` 패턴, 명사형 종결("~확대", "~개선") `[공식]`
- **러닝헤드 + 페이지번호**: 우상단 행사명, 하단 페이지번호+로고 `[공식]`
- **각주**: ※ 기호로 표·차트 하단에 `[공식]`

### 1.3 분량 — 10/20/30 법칙과 10% 규칙

- **Guy Kawasaki 10/20/30** (원전 직접 확인): 슬라이드 10장 / 발표 20분 / 폰트 30pt 미만 금지. 근거: "보통 사람은 한 미팅에서 10개 이상의 개념을 이해하지 못한다." 1시간 슬롯이어도 20분에 끝내고 나머지는 토론. — https://guykawasaki.com/the_102030_rule/ `[공식]`
- **Duarte 10% 규칙**: 자료가 50장이면 요약 5장으로 발표하고 나머지는 부록(질문 대응용). "슬롯이 5분으로 잘렸다고 가정"하고 결론·권고를 맨 앞에. "임원은 마지막에 반전이 있는 긴 발표를 견디지 않는다 — 끝나기 전에 끊는다"(HBR, Nancy Duarte 2012). — https://www.duarte.com/blog/how-to-effectively-present-to-senior-executives/ `[공식]`
- 한국 공공부문도 "1건 1매" 경제성 원칙을 명문화 (경기도교육청 공문서 작성법, 행정업무규정 근거). — https://www.goe.go.kr/resource/old/BBSMSTR_000000000028/BBS_202410150153084250.pdf `[공식]`

---

## 2. 스토리 설계 — 결론 우선 (두괄식)

### 2.1 피라미드 원칙 (Barbara Minto)

- 생각을 **하나의 포인트 아래 피라미드로 조직**해 제시. 사고는 아래→위로 하되 **전달은 위→아래(결론 먼저)**. — https://www.barbaraminto.com/ `[공식]`
- 3대 구조 규칙: ① 상위 포인트는 하위 그룹의 요약 ② 그룹 내 아이디어는 논리적 동종 ③ 그룹은 MECE(상호배타·전체포괄). MECE는 Minto 본인의 발명. `[미확인 — 2차 요약, 단 MECE 발명은 McKinsey 공식 인터뷰로 확인]`

### 2.2 도입부: SCQ 프레임워크

Minto 공식 사이트가 명시한 "독자 머릿속의 질문을 찾는" 도입 구조. `[공식]`

1. **Situation**: 청중 전원이 동의하는 사실로 시작 (논쟁 없는 출발점)
2. **Complication**: 상황을 흔드는 문제·기회 → 청중 머릿속에 자연스럽게 질문이 생김
3. **Question → Answer**: 그 질문에 대한 답 = 이 보고서의 핵심 권고

### 2.3 행동 제목 (Action Title)

- 제목 = 슬라이드의 가장 중요한 포인트를 **완결 문장**으로. **최대 15단어, 2줄 초과 금지**. — https://slideworks.io/resources/how-to-write-action-titles-like-mckinsey `[공식]`
- 수치·동인을 제목에 포함: "공급망 최적화 가능"(약함) → "공급망 프로세스 최적화로 비용 20% 절감"(강함)
- **수평 논리 테스트**: 제목만 순서대로 읽어도 전체 스토리가 완성되어야 함
- 학술 근거(Assertion-Evidence, Penn State Michael Alley): 제목은 구(phrase)가 아닌 주장(assertion), 본문은 불릿이 아닌 시각적 증거. — https://www.assertion-evidence.com/ `[공식]`

### 2.4 Executive Summary 작성법

- 구성: Situation → Complication → Resolution (+ Call to Action). **1장 원칙(최대 2장)**, 위치는 표지 직후·목차보다 앞.
- 포맷: **볼드-불릿 구조** — 핵심 테이크어웨이를 볼드 문장으로, 그 아래 수치·근거 불릿. 볼드만 읽어도 스토리가 완성되어야 함. — https://slideworks.io/resources/how-to-write-executive-summary `[공식]`

### 2.5 발표용 vs 전달용(읽기용) 구분

- Duarte Slidedoc 원칙: **용도가 다르면 덱도 달라야 한다**. 발표용은 시각 중심·텍스트 최소, 전달용(leave-behind)은 단독으로 읽히는 문서. 듀얼 용도가 필요하면 본문을 빽빽하게 만들지 말고 **Notes 영역을 활용**해 읽기용을 별도 생성. — https://www.duarte.com/blog/the-slides-you-deliver-versus-the-slidedoc-you-leave-behind/ `[공식]`
- 한국 실무 보정: 고객사 전달 PPT는 사실상 전달용(읽기용) 성격이 강함 → 발표 없이 읽혀도 이해되도록 헤드라인·불릿의 자기완결성을 높이되, 깨알 표는 별첨으로 격리 (삼성 패턴과 일치) `[공식 — 1.1 근거]`

---

## 3. 타이포그래피·레이아웃 — Microsoft 공식 수치

### 3.1 슬라이드 규격

- **16:9 Widescreen, 13.333 × 7.5 in (33.867 × 19.05 cm)** — Microsoft 공식 "best choice", 신규 기본값. — https://support.microsoft.com/en-us/office/change-the-size-of-your-powerpoint-slides-040a811c-be43-40b9-8d04-0de5ed79987e `[공식]`
- 여백: **콘텐츠는 슬라이드 가장자리까지 확장 금지** (Microsoft Designer 템플릿 공식 규칙 — Picture placeholder만 가장자리 허용). — https://support.microsoft.com/en-us/office/creating-custom-templates-that-work-well-with-designer-in-powerpoint-21521084-0c21-4471-bec1-a286a2f70b9f `[공식]`
- "Safe zone"이라는 용어의 PowerPoint 공식 문서는 없음 — 위 여백 규칙이 가장 근접한 공식 권고 `[공식 — 부재 확인]`

### 3.2 글꼴 크기

| 용도 | 크기 | 근거 |
|------|------|------|
| 절대 하한 | **18pt 이상** | Microsoft 디자인 가이드·접근성 가이드 양쪽 일치 `[공식]` |
| 발표용 본문 권장 | 30pt 내외 | Kawasaki 30pt 법칙 `[공식]` (Microsoft 문서의 30pt 언급은 `[미확인]`) |
| 전달용(읽기용) 본문 | 18pt를 하한으로 용도에 맞게 | Slidedoc은 읽기 문서이므로 발표용보다 작게 허용 `[공식 — Duarte]` |

- 불릿: **항목당 1줄, 줄바꿈 없이**. 관사·군더더기 제거로 단어 수 축소. — https://support.microsoft.com/en-us/office/tips-for-creating-and-delivering-an-effective-presentation-f43156b0-20d2-4c51-8345-0c337cefb88b `[공식]`
- 전체 대문자·과도한 이탤릭/밑줄 금지, 산세리프 사용 `[공식 — MS 접근성 가이드]`

### 3.3 한국어 폰트 선택 — 라이선스가 핵심

| 폰트 | 라이선스 | 고객사 전달 시 |
|------|---------|---------------|
| **맑은 고딕** | Microsoft 독점(무료 폰트 아님). MS 제품 내 사용만 허용 | Office 문서 교환 맥락에선 통상 문제없으나, 폰트 파일 추출·재배포·임베딩은 라이선스 리스크. — https://learn.microsoft.com/ko-kr/answers/questions/2655428/question-2655428 `[공식]` |
| **Pretendard** | SIL OFL — 임베딩·재배포·상업 사용 모두 허용 | **고객사 전달 PPT에 법적으로 가장 안전**. — https://github.com/orioncactus/pretendard `[공식]` |
| 노토산스 KR | SIL OFL (Google Fonts) | 임베딩 허용 `[미확인 — 배포 페이지 직접 미열람]` |

- 맑은 고딕은 ClearType 기반으로 "화면 가독성 우수, 작은 크기에서도 매우 읽기 좋음"이 Microsoft 공식 문서에 명시 — 수신 환경이 Windows/Office로 보장되면 합리적 선택. — https://learn.microsoft.com/en-us/typography/font-list/malgun-gothic `[공식]`
- **실무 권고**: 사내·Windows 환경 확정 → 맑은 고딕 / 고객사 전달·환경 불확실 → Pretendard 임베딩

---

## 4. 색상 — 단일 강조색 + 접근성 수치

### 4.1 색 운용 원칙 (한국 대기업 실물 패턴)

- **강조색은 브랜드 단일색 하나만** (삼성 블루, 카카오 옐로). 당기/핵심 데이터만 강조색, 비교 기간·보조 데이터는 회색. `[공식 — 1.1 근거]`
- 배경: 표지·간지는 어두운 브랜드색+흰 글자, 데이터 장은 흰 배경. `[공식 — 삼성 패턴]`
- 색에만 의존 금지 — 색+텍스트, 색+모양 병행(인구 약 15%가 색각 관련 어려움). — Microsoft 접근성 영상 문서 `[공식]`

### 4.2 대비 수치 (WCAG 2.2, W3C 직접 확인)

| 대상 | 기준 (Level AA) | 출처 |
|------|----------------|------|
| 일반 텍스트 | **4.5:1** 이상 | WCAG 2.2 SC 1.4.3 `[공식]` |
| 큰 텍스트 (18pt 이상 또는 14pt 볼드) | **3:1** 이상 | 동일 `[공식]` |
| 차트 막대·선·범례 등 그래픽 객체 | 인접 색과 **3:1** 이상 | SC 1.4.11 Non-text Contrast `[공식]` |

— https://www.w3.org/TR/WCAG22/ `[공식]`. PowerPoint 자체 문서에는 수치가 없고 Accessibility Checker 사용 권고만 있음 `[공식 — 부재 확인]`.

### 4.3 색각이상 대응 팔레트 (Okabe-Ito 8색)

차트 시리즈 색상의 사실상 표준. 원전: https://jfly.uni-koeln.de/color/ `[공식]`, 수치: Wilke, *Fundamentals of Data Visualization* — https://clauswilke.com/dataviz/color-pitfalls.html `[공식]`

```
orange  #E69F00   sky blue       #56B4E9   bluish green  #009E73   yellow  #F0E442
blue    #0072B2   vermilion      #D55E00   reddish purple #CC79A7   black   #000000
```

- 순수 빨강 대신 vermilion, 보라 대신 reddish purple 사용
- **색+모양/패턴 이중 부호화(redundant coding)** 필수

---

## 5. 데이터 시각화 — 차트·표

### 5.1 차트 유형 선택 (FT Visual Vocabulary)

"데이터의 관계 유형 → 차트" 매핑. — https://github.com/Financial-Times/chart-doctor/blob/main/visual-vocabulary/README.md `[공식]`

| 보여줄 관계 | 권장 차트 |
|------------|----------|
| 시간 변화 | 라인, 컬럼, 영역 |
| 크기 비교 | 막대(세로/가로) |
| 순위 | 정렬된 막대, 슬로프, 롤리팝 |
| 구성(part-to-whole) | 누적 막대, 트리맵, 워터폴 |
| 편차(기준 대비) | 다이버징 바 |
| 상관 | 산점도, 버블 |
| 분포 | 히스토그램, 박스플롯 |

- **파이차트**: "정확한 비교를 어렵게 한다" — 전면 금지는 아니나 신중 사용(항목 적고 구성비 강조일 때만) `[공식]`

### 5.2 IBCS SUCCESS 원칙 (보고서 차트의 국제 표준)

IBCS v1.2, CC BY-SA 무료 공개. — https://www.ibcs.com/standards/ `[공식]`

- **SAY** 메시지를 말하라 · **UNIFY** 동일 의미는 동일 표기 · **CONDENSE** 정보 밀도를 높여라 · **CHECK** 시각적 무결성(축 절단 금지) · **EXPRESS** 적절한 시각화 선택 · **SIMPLIFY** 군더더기 제거 · **STRUCTURE** 콘텐츠 구조화
- 한국 IR 실물과의 합치: 카카오·삼성 모두 "당기=진한 색, 과거=회색" — IBCS UNIFY(의미적 표기 통일)와 일치 `[공식]`

### 5.3 표 디자인 (Schwabish 10규칙)

Schwabish, *Journal of Benefit-Cost Analysis* 11(2), 2020. — https://doi.org/10.1017/bca.2020.11 `[공식 — 초록]`, 규칙 목록 `[공식 — 2차 출처]`

1. 숫자 우측 정렬 2. 텍스트 좌측 정렬 3. 소수점 1~2자리 4. 단위($, %)는 첫 행에만 5. 이상치 강조 6. 열 헤더 강조 7. 구분선·강조 절제 8. 행·열 여백 확보 9. 그룹 구분은 여백/구분선 10. 큰 표는 히트맵 등 시각화로 대체
- 한국 실물 보정: 상세 표는 본문이 아닌 **별첨으로 격리**가 대기업 표준 (삼성 패턴) `[공식]`
- Microsoft 접근성: 표는 가능하면 회피, 사용 시 머리글 행 지정·셀 병합 금지 `[공식]`

---

## 6. 2026년 트렌드 (참고 — 벤더 데이터 기반, 보수적 적용)

SlideEgg 자사 10만+ 다운로드 분석. — https://www.slideegg.com/blog/presentation-tips/the-7-presentation-design-trends-dominating-2026-data-backed/ `[공식 — 단 벤더 블로그]`

- **Bento Grid 레이아웃**: 카드형 분할 구성 — 한국 IT 보고서에 무난히 수용 가능
- **초대형 타이포그래피**: 핵심 수치를 화면의 큰 비중으로 — 삼성 요약 장의 "대형 KPI 숫자" 패턴과 합치, 적극 권장
- **Data Storytelling**: 엑셀 스크린샷 대신 단일 인사이트 차트 — IBCS SAY 원칙과 합치
- 다크모드·글래스모피즘·세로형(9:16): 임원·고객사 보고 맥락에선 보수적으로 — 표지·간지 한정 적용 권장 `[추측 — 트렌드의 보고서 적용 범위는 본 문서의 판단]`

Canva 2026 트렌드 리포트는 직접 열람 실패 `[미확인]`.

---

## 7. 체크리스트 — 납품 전 최종 점검

### 구조
- [ ] 표지 직후에 요약 1장(Executive Summary)이 있는가 — 볼드만 읽어도 스토리가 되는가
- [ ] 제목만 순서대로 읽어도 전체 논리가 이어지는가 (수평 논리 테스트)
- [ ] 모든 헤드라인이 결론(가급적 수치)으로 시작하는가
- [ ] 상세 표·깨알 데이터가 본문에 남아 있지 않은가 (→ 별첨 격리)
- [ ] 외부 전달 자료라면 면책/기준 고지가 있는가

### 타이포·레이아웃
- [ ] 18pt 미만 텍스트가 없는가 (발표용은 30pt 내외)
- [ ] 불릿이 항목당 1줄을 넘지 않는가
- [ ] 콘텐츠가 슬라이드 가장자리에 붙어 있지 않은가
- [ ] 고객사 전달본이면 폰트 라이선스 확인(맑은 고딕 임베딩 리스크 → Pretendard 검토)
- [ ] 모든 슬라이드에 고유 제목이 있는가 (접근성)

### 색·차트
- [ ] 강조색이 단일색인가 — 핵심 데이터만 강조색, 비교군은 회색인가
- [ ] 텍스트 대비 4.5:1 / 큰 텍스트·그래픽 3:1을 충족하는가
- [ ] 색에만 의존한 구분이 없는가 (색+라벨/모양 병행)
- [ ] 차트 유형이 데이터 관계와 맞는가 (시간=라인, 비교=막대, 구성=누적/트리맵)
- [ ] 축 절단으로 왜곡된 차트가 없는가 (IBCS CHECK)
- [ ] 표: 숫자 우측 정렬, 단위는 첫 행만, 구분선 절제

---

## 출처 전체 목록

### Microsoft 공식
- [Tips for creating and delivering an effective presentation](https://support.microsoft.com/en-us/office/tips-for-creating-and-delivering-an-effective-presentation-f43156b0-20d2-4c51-8345-0c337cefb88b)
- [Make your PowerPoint presentations accessible](https://support.microsoft.com/en-us/office/make-your-powerpoint-presentations-accessible-to-people-with-disabilities-6f7772b2-2f33-4bd2-8ca7-dae3b2b3ef25)
- [Change the size of your PowerPoint slides](https://support.microsoft.com/en-us/office/change-the-size-of-your-powerpoint-slides-040a811c-be43-40b9-8d04-0de5ed79987e)
- [Creating custom templates that work well with Designer](https://support.microsoft.com/en-us/office/creating-custom-templates-that-work-well-with-designer-in-powerpoint-21521084-0c21-4471-bec1-a286a2f70b9f)
- [Malgun Gothic font family](https://learn.microsoft.com/en-us/typography/font-list/malgun-gothic)
- [Microsoft Q&A — 맑은 고딕 사용 범위](https://learn.microsoft.com/ko-kr/answers/questions/2655428/question-2655428)

### 한국 실물·표준
- [삼성전자 2025 1Q 경영설명회](https://images.samsung.com/kdp/ir/events/2025/2025_1Q_conference_kor.pdf)
- [카카오 3Q25 실적발표](https://t1.kakaocdn.net/kakaocorp/admin/ir/results-announcement/5894.pdf)
- [경기도교육청 공문서 작성법(행정업무규정 기반)](https://www.goe.go.kr/resource/old/BBSMSTR_000000000028/BBS_202410150153084250.pdf)
- [국가기술표준원 KS 색채표준](https://www.kats.go.kr/content.do?cmsid=83)
- [Pretendard (SIL OFL)](https://github.com/orioncactus/pretendard)

### 프레젠테이션 설계 원전
- [Guy Kawasaki — 10/20/30 Rule (원전)](https://guykawasaki.com/the_102030_rule/)
- [Barbara Minto 공식 사이트](https://www.barbaraminto.com/)
- [Duarte — Slidedoc vs 발표 덱](https://www.duarte.com/blog/the-slides-you-deliver-versus-the-slidedoc-you-leave-behind/)
- [Duarte — 임원 대상 발표법 (HBR 2012 연계)](https://www.duarte.com/blog/how-to-effectively-present-to-senior-executives/)
- [Slideworks — Action Titles](https://slideworks.io/resources/how-to-write-action-titles-like-mckinsey) · [Executive Summary](https://slideworks.io/resources/how-to-write-executive-summary)
- [Assertion-Evidence (Penn State)](https://www.assertion-evidence.com/)

### 데이터 시각화·접근성
- [W3C WCAG 2.2](https://www.w3.org/TR/WCAG22/)
- [FT Visual Vocabulary](https://github.com/Financial-Times/chart-doctor/blob/main/visual-vocabulary/README.md)
- [IBCS Standards v1.2](https://www.ibcs.com/standards/)
- [Okabe & Ito — Color Universal Design (원전)](https://jfly.uni-koeln.de/color/)
- [Wilke — Fundamentals of Data Visualization Ch.19](https://clauswilke.com/dataviz/color-pitfalls.html)
- [Schwabish — Ten Guidelines for Better Tables](https://doi.org/10.1017/bca.2020.11)

### 트렌드 (참고)
- [SlideEgg — 2026 Presentation Design Trends](https://www.slideegg.com/blog/presentation-tips/the-7-presentation-design-trends-dominating-2026-data-backed/)
