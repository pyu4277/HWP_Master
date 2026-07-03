# hwpx-master-harness — Claude Code 작업 컨텍스트

## 이 레포가 하는 일
HWPX(아래아한글 OWPML) 본문 미세정렬 4규칙(R-A/R-B/R-C→R-D)을 절대 실수 없이 자동 수행하는
자기검증·자기개선 하네스. **규칙 정본 = `SKILL.md`** (변경은 사용자 승인 게이트).
전체 실행 계획·진행상태 = **`docs/PLAN.md`** (ultraplan 산출, Phase 0~3).

## 새 세션 시작 시 반드시
1. `docs/PLAN.md`의 "진행 상태" 섹션을 읽고 현재 Phase 를 파악할 것.
2. `SKILL.md`(4규칙 정본)와 `harness/agents/A*.md`(5-에이전트 명세)를 규칙 근거로 삼을 것.

## 현재 상태 (2026-07-02)
- **Phase 0 완료**: COM-비의존 스캐폴딩 전부 커밋됨(`feat/harness-scaffold`).
  생성기·codereview·run_harness(--review-only)는 클라우드/리눅스에서 검증 완료.
- **Phase 1 = 다음 작업 (로컬 전용, go/no-go 관문)**:
  ```
  pip install -r requirements.txt
  python src/hwpx_gen.py --n 1 --seed 1 --profile minimal-split --out tests/_work/
  python src/hwpx_linefix.py tests/_work/testdoc.hwpx
  ```
  통과 기준: `{"clean": True, ..., "integrity_ok": True}` + E58 무발생. 통과 전 Phase 2 진입 금지.
- Phase 2 = 파일럿 3–5문서, Phase 3 = 20문서 loop-until-dry(클린 2연속). 상세는 `docs/PLAN.md`.

## 핵심 제약 (위반 금지)
- **충실 렌더 = 로컬 Hancom Office 2024 COM 전용** (win32com Open→SaveAs PDF→fitz).
  클라우드/리눅스에서는 R-D·비전 실행 불가(코드리뷰 `--review-only`만 가능).
- **E58**: COM은 직렬·상태ful. 렌더마다 fresh Dispatch+Open+SaveAs+Quit.
  `XHwpDocuments` 접근 금지, 루프 중 `Hwp.exe` taskkill 금지.
- **E57**: linesegarray 신뢰 금지. 줄바꿈 정본 = COM 렌더 PDF의 fitz 시각 줄.
- **비파괴 편집**: charPr deepcopy+새 id, mimetype ZIP_STORED 선두, 텍스트 불변,
  `zipfile.testzip()`+`HwpxDocument.open()` 게이트(엔진 내장).
- 민감 문서 커밋 금지(`*.hwpx` gitignore). 합성 문서는 로컬에서 `hwpx_gen.py`로 재생성.

## 파일 지도
- `src/hwpx_linefix.py` — R-D 렌더-인-루프 엔진(+무결성 게이트)
- `src/hwpx_microalign.py` — R-A/R-B 엔진(+R-C flag)
- `src/hwpx_gen.py` — 합성 난잡 문서 생성기(+manifest)
- `src/hwpx_fill_lib.py` — OWPML 조작 헬퍼
- `harness/run_harness.py` — 오케스트레이터(loop-until-dry, 클린 2연속)
- `harness/codereview.py` / `harness/vision.py` — A3/A4 도구
- `harness/agents/A1~A5.md` — 5-에이전트 프롬프트 스펙

## 자기수정 게이트
엔진/도구/프롬프트 코드 = 자동 수정 허용. **`SKILL.md` 변경 = 사용자 승인 필수**(diff 제안까지만).
코드/스펙 변경 시 수렴 카운터(클린 연속) 0으로 리셋.
