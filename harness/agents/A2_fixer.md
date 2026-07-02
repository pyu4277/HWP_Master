# A2 — 독립 수정 (Fixer)

**역할**: 생성 에이전트와 **분리**되어, 결과 HWPX 만 보고 4규칙대로 수정한다(manifest 비공개).
도구 = `src/hwpx_microalign.py`(R-A/R-B), `src/hwpx_linefix.py`(R-D 렌더-루프, **로컬 COM**).

**입력**: `docNN.hwpx` (기대치 비공개).
**출력**: `docNN.fixed.hwpx` + change-log(규칙별 적용 내역, 각 도구의 stdout).

**처리 순서**:
1. `hwpx_microalign.py <src> <fixed> --rules AB` → R-A(음수 intent 정규화)·R-B(선두 괄호 bold).
2. `hwpx_linefix.py <fixed>` → R-D 렌더-인-루프(단절 0까지). **로컬 Windows + Hancom COM 필요**.

**검증기준(자체)**: 텍스트 불변 + `zipfile.testzip()` + `HwpxDocument.open()` PASS(도구 내장 게이트),
**멱등성**(재실행 시 추가 변경 0). 실패 시 A5 로 에스컬레이션.
