# A4 — 비전 검증 (Vision)

**역할**: 수정 HWPX 를 COM 으로 렌더한 PNG 를 비전으로 판정. **렌더 정본**(linesegarray/manifest 보다 우선).
도구 = `harness/vision.py`. **로컬 전용(COM)**.

**입력**: `docNN.fixed.hwpx`.
**처리**: `vision.split_report()`(COM→PDF→fitz 단절 실측) + `vision.render_pngs()`(페이지별 PNG).
**출력**: 페이지별 판정 + 실패 주석. 객관 백본 = `{count(단절수), splits[], pages}`.

**검증기준(렌더 실측)**:
- 단절 어절 **0**(육안 + `detect_splits`).
- 자동 글머리 접힌 줄이 첫줄 본문에 정렬(R-A).
- 본문 선두 괄호 주제어가 진하게(R-B), 제목/캡션/표헤더는 미변경.
- 배제 대상(양수 intent·수동 글머리) 미변경.

`count>0` 또는 육안 결함 → DIRTY → A5 입력. 백본 단절수 0 이 클린의 **필요조건**, 비전 육안이 최종 확정.
