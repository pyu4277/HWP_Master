# hwpx-master-harness

HWPX(아래아한글 OWPML) 본문 미세정렬 규칙을 **절대 실수 없이 자동 수행**하는 self-verifying 하네스.
사용자 한/글 완성본을 ground truth 로, 규칙별 실현성을 실측 + 적대검증한 뒤 승격한다.

이 레포는 **계획(ultraplan) + 버전관리용**이다. 민감 문서는 포함하지 않는다(`*.hwpx` gitignore).
**실행은 로컬 전용** — 충실 렌더가 로컬 Hancom Office 2024 COM 에 의존하기 때문(아래 제약).

## 본문 미세정렬 4규칙

- **R-A 글머리 내어쓰기 정렬 (FULL, 헤드리스)** — 자동 글머리(paraPr `heading type=BULLET`) 문단의 음수 첫줄 내어쓰기(`intent`)만 0 으로 통일해 접힌 줄이 첫줄 본문 아래 정렬. 양수 들여쓰기·수동 글머리·제목은 불가침. 그룹 left(왼쪽여백)는 보존.
- **R-B 본문 선두 괄호 주제어 진하게 (FULL, 헤드리스)** — 문단 선두 `(...)/[...]/<...>` 주제어 토큰만 bold. 게이트: 선두 run 스타일(height,fontRef)이 본문 최빈값과 일치 + 토큰 뒤 본문 잔여 >= 5자. 제목/표헤더/캡션 배제. 토큰을 독립 run 으로 분리해 bold 쌍둥이 charPr(크기·글꼴 동일, `bold` 만 on) 지정. 텍스트 불변.
- **R-C 자간 2줄->1줄 (R-D로 승격)** — 렌더 개방 전엔 flag-only(후보 제시). 이제 R-D 로 흡수.
- **R-D 정밀 자간 단절방지 (FULL, COM 렌더-인-루프)** — 아래 상세.

### R-D 상세 (핵심 신규)

어절(공백 구분 토큰)이 줄 끝에서 단절(한글 기본 글자단위 줄나눔으로 어절 중간이 다음 줄로 넘어감)되면 자간으로 **무조건 1줄화**.

- **방향 최적화**: 단절 어절의 head(현재 줄 잔여 글자수) vs tail(다음 줄 글자수) 비교 -> 옮길 글자 수가 적은 쪽. head<tail 이면 widen(head 를 아래로 밀어 tail 과 합침, 양수 자간), head>=tail 이면 narrow(tail 을 위로 당김, 음수 자간). 예: "대학이"(head 1 < tail 2)->widen; "요구되는"(head 1 < tail 3)->widen.
- **자간 범위(필수)**: 해당 줄 첫 글자 ~ 조절 마지막 어절까지의 셀만. 문단/문장 전체 금지. **선두 글머리 기호 제외**(글머리 뒤 정렬 열이 자간에 밀려 삐뚤어짐 방지 = R-A 보존).
- **연쇄 파급 처리**: 한 곳 자간 변경이 하류 줄바꿈을 바꾸므로 위->아래 순차 + 매 조절 후 재렌더로 실측(추정 아님). 무단절될 때까지 반복(loop-until-clean).
- **자간 크기**: 기본 하한 -20%(육안 예외 허용), widen 동일 범위. 최소 자간으로 단절 해소되는 값을 증분/이분 탐색.

## 핵심 제약 (반드시 준수)

- **충실 렌더 = 로컬 Hancom Office 2024 COM 전용.** python `win32com.client.Dispatch("HWPFrame.HwpObject")` -> `Open(abs,"HWPX","")` -> `SaveAs(absPdf,"PDF","")` -> PDF 를 PyMuPDF(`fitz`)로 PNG/텍스트 추출. 클라우드/원격엔 한/글 부재 -> **계획만 가능, 실행 불가**.
- **비파괴 편집**: 공유 charPr 는 deepcopy 후 새 id, mimetype 은 ZIP_STORED 선두 repack, 편집 후 텍스트 불변(문단 텍스트 연결 동일) 회귀검사 + `hwpx.HwpxDocument.open()` + `zipfile.testzip()` 무결성.
- **linesegarray 신뢰 금지**: python-hwpx 편집이 제거하거나 COM SaveAs 가 leaf 문단에 재생성 안 함 -> stale/부재. **줄바꿈 정본 = COM 렌더 PDF**(fitz `get_text("dict")` 시각 줄).
- COM 은 직렬·상태ful -> 렌더 순차. 대량 문서는 백그라운드 순차.

## 실측 증명된 primitive (2026-07-02)

1. COM 렌더 = python win32com Open(HWPX) -> SaveAs(PDF) 성공.
2. 단절 어절 탐지 = fitz 시각 줄에서 "줄끝 한글 + 다음줄첫 한글 + 경계 공백 없음".
3. 자간 쓰기 = lxml `<hh:spacing hangul>` 부분 run 분리.
4. narrow/widen 효율 로직(사용자 예시와 정확 일치).

## 코드 상태

규칙 정본 스펙은 [`SKILL.md`](SKILL.md)로 이관(단일 근거). 이 README 는 배경·제약 요약.

- `src/hwpx_microalign.py` — R-A/R-B(+R-C flag) 검증 완료(독립 verifier: ground truth 15/15 일치, 멱등성·텍스트불변·무결성 PASS).
- `src/hwpx_fill_lib.py` — 재사용 OWPML 헬퍼(load_section/serialize/repack/run·paraPr 조작).
- `src/hwpx_linefix.py` — **R-D 렌더-인-루프 엔진 구현 완료**. fresh-Dispatch 개정판 완주 검증 대기(E58 관문, 로컬 전용). 텍스트불변·testzip·hwpx-open 무결성 게이트 포함.
- `src/hwpx_gen.py` — 표·캡션·주석 포함 난잡 합성 문서 생성기(+manifest).
- `harness/` — 5 에이전트 자가개선 하네스(생성/독립수정/코드리뷰/비전/자기수정, loop-until-dry).

## 자가개선 하네스 (구축 대상)

5 에이전트 파이프라인, 클린 2연속까지 loop-until-dry:
1. **생성** — 표/캡션/주석 포함 최대한 복잡·난잡한 예시 HWPX 20개 생성(합성, 민감데이터 없음).
2. **독립 수정** — 생성 에이전트와 분리, 결과 HWPX 만 보고 4규칙대로 수정.
3. **코드리뷰 검증** — 논리/코드/OWPML 관점 수정 정합성 검토.
4. **비전 검증** — COM 렌더 PNG 를 비전으로 판단(단절 0, 정렬·bold 정확).
5. **자기수정** — 틀린 원인 분석 -> 지침/도구/엔진 보완(self-update). 전 문서 연속 2라운드 무결점까지 반복.

## 수용 기준

- [ ] R-D 렌더-인-루프 엔진: 단절 탐지 -> 방향최적 자간(줄-스팬, 글머리 제외) -> 재렌더 반복 -> 단절 0. 텍스트불변·무결성.
- [ ] 실제/합성 문서에서 단절 어절이 렌더상 사라짐을 PDF->PNG 비전으로 확인.
- [ ] 20 합성 난잡문서 전부가 코드리뷰+비전 이중검증 통과(클린 2연속).
- [ ] 지침(SKILL.md 4규칙)·도구가 하네스로 수렴·확정.

## 로컬 실행 전제

Windows + Hancom Office 2024(COM) + python 3.12(`win32com`, `lxml`, `hwpx`(python-hwpx), `PyMuPDF`(fitz)).
클라우드 ultraplan 은 계획 산출용이며, 산출 플랜을 로컬로 가져와 위 전제에서 실행한다.
