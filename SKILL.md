---
name: hwpx-microalign
description: >-
  HWPX(아래아한글 OWPML) 본문 미세정렬 4규칙을 절대 실수 없이 자동 수행한다.
  R-A 글머리 내어쓰기 정렬 / R-B 본문 선두 괄호 주제어 진하게 / R-C 자간 2줄→1줄(R-D로 흡수) /
  R-D 정밀 자간 단절방지(COM 렌더-인-루프). 충실 렌더는 로컬 Hancom Office 2024 COM 전용.
  이 파일이 규칙의 단일 정본(single source of truth)이며, 변경은 사용자 승인 게이트 대상이다.
---

# HWPX 본문 미세정렬 4규칙 (정본 스펙)

> 이 문서는 규칙의 **정본**이다. `README.md`·`src/*`·`harness/*` 는 이 스펙을 구현·검증한다.
> 자가개선 하네스(A5 자기수정)가 이 파일 변경을 **제안**할 수 있으나, **적용은 사용자 승인 후에만** 한다.
> ground truth = 사용자 한/글 완성본. 규칙은 실측 + 적대검증으로 승격됐다.

## 핵심 원칙 (반드시 준수)

- **비파괴 편집**: 공유 charPr 는 `deepcopy` 후 **새 id** 부여(원본 불변). mimetype 은 `ZIP_STORED`
  로 zip 선두에 repack. 편집 후 **텍스트 불변**(문단 텍스트 연결 동일) 회귀검사 + `zipfile.testzip()`
  + `hwpx.HwpxDocument.open()` 무결성 검사를 **강제**하고, 실패 시 롤백한다.
- **linesegarray 신뢰 금지(E57)**: python-hwpx 편집이 제거하거나 COM SaveAs 가 leaf 문단에 재생성하지
  않아 stale/부재일 수 있다. **줄바꿈 정본 = COM 렌더 PDF**(fitz `get_text("dict")` 시각 줄).
- **충실 렌더 = 로컬 Hancom Office 2024 COM 전용**: `win32com.client.Dispatch("HWPFrame.HwpObject")`
  → `Open(abs,"HWPX","")` → `SaveAs(absPdf,"PDF","")` → PyMuPDF(`fitz`)로 시각 줄/PNG 추출.
  클라우드/원격엔 한/글 부재 → 계획·COM-비의존 코드만 가능, 렌더 실행은 로컬.
- **COM 직렬·상태ful(E58)**: 렌더마다 **fresh `Dispatch`+`Open`+`SaveAs`+`Quit`**. `XHwpDocuments`
  **미접근**. 루프 중 `Hwp.exe` **taskkill 금지**(COM 서버 오염 → "RPC 서버 사용 불가" → `AttributeError('Open')`).
  렌더는 순차. 대량 문서는 백그라운드 순차 큐.

---

## R-A — 글머리 내어쓰기 정렬 (FULL, 헤드리스)

자동 글머리(paraPr `heading type="BULLET"`) 문단의 **음수 첫줄 내어쓰기(`intent`)만 0 으로 통일**해,
접힌(둘째 줄 이후) 줄이 첫줄 본문 아래로 정렬되게 한다.

- **대상**: `heading@type="BULLET"` 이고 **≥2 멤버**인 실제 리스트 그룹의 paraPr, 그 안의 음수 `margin/intent`.
- **불가침**: 양수 들여쓰기, 수동 글머리, 제목. 그룹 `left`(왼쪽 여백)는 **보존**.
- **동작**: 음수 `intent@value` → `0`. 텍스트·다른 속성 불변.
- 구현: `src/hwpx_microalign.py::rule_a`.

## R-B — 본문 선두 괄호 주제어 진하게 (FULL, 헤드리스)

문단 선두의 `(...)` / `[...]` / `<...>` **주제어 토큰만** bold 처리한다.

- **게이트**: (a) 선두 run 스타일(height, fontRef)이 **본문 최빈값과 일치**, (b) 토큰 뒤 본문 잔여 **≥ 5자**.
- **배제**: 제목/표헤더/캡션(선두 스타일이 본문 최빈과 다르면 배제), 이미 bold 인 선두(멱등성).
- **동작**: 토큰을 **독립 run 으로 분리**, 크기·글꼴 동일하고 `bold` 만 on 인 **쌍둥이 charPr**(새 id) 지정.
  **텍스트 불변**.
- 구현: `src/hwpx_microalign.py::rule_b` (+ `get_bold_twin`, `single_t_run`, `body_mode_style`).

## R-C — 자간 2줄→1줄 (R-D 로 승격/흡수)

렌더 개방 전에는 **flag-only**(보수적 후보만 제시, 자동 붕괴 금지). 렌더 접근 후에는 **R-D 로 흡수**되어
실제 자간 조절은 R-D 엔진이 렌더 실측으로 수행한다.

- flag 구현(참고): `src/hwpx_microalign.py::rule_c_flag`.

## R-D — 정밀 자간 단절방지 (FULL, COM 렌더-인-루프)

어절(공백 구분 토큰)이 줄 끝에서 **단절**(한글 기본 글자단위 줄나눔으로 어절 중간이 다음 줄로 넘어감)되면
**자간으로 무조건 1줄화**한다.

- **탐지**: fitz 시각 줄에서 "줄끝 한글 + 다음줄첫 한글 + 경계 공백 없음"(`detect_splits`).
- **방향 최적화**: 단절 어절의 head(현재 줄 잔여 글자수) vs tail(다음 줄 글자수) 비교 → 옮길 글자 수가
  적은 쪽. `head < tail` → **widen**(head 를 아래로 밀어 tail 과 합침, 양수 자간);
  `head ≥ tail` → **narrow**(tail 을 위로 당김, 음수 자간). 예: "대학이"(head1<tail2)→widen,
  "요구되는"(head1<tail3)→widen.
- **자간 범위(필수)**: 해당 줄 **첫 글자 ~ 조절 마지막 어절**까지의 셀만. 문단/문장 전체 금지.
  **선두 글머리 기호 제외**(글머리 뒤 정렬 열이 자간에 밀려 삐뚤어짐 방지 = R-A 보존; `skip_bullet`).
- **연쇄 파급 처리**: 한 곳 자간 변경이 하류 줄바꿈을 바꾸므로 **위→아래 순차 + 매 조절 후 재렌더로 실측**
  (추정 아님). 무단절될 때까지 반복(**loop-until-clean**).
- **자간 크기**: 기본 하한 −20%(육안 예외 허용), widen 동일 범위. 최소 자간으로 단절 해소되는 값을
  증분/이분 탐색.
- 구현: `src/hwpx_linefix.py`(`render_pdf`/`visual_lines`/`detect_splits`/`map_split`/`apply_span_spacing`/
  `spacing_twin`/`skip_bullet`/`fix_document`).

---

## 수용 기준

- [ ] R-D 렌더-인-루프: 단절 탐지 → 방향최적 자간(줄-스팬, 글머리 제외) → 재렌더 반복 → 단절 0. 텍스트불변·무결성.
- [ ] 실제/합성 문서에서 단절 어절이 렌더상 사라짐을 PDF→PNG 비전으로 확인.
- [ ] 20 합성 난잡문서 전부가 코드리뷰+비전 이중검증 통과(**클린 2연속**).
- [ ] 이 SKILL.md(4규칙)·도구가 하네스로 수렴·확정.
