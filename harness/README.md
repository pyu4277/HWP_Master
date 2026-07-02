# harness/ — 5-에이전트 자기개선 하네스

규칙 정본 = [`../SKILL.md`](../SKILL.md). 이 하네스는 4규칙을 실측·적대검증해 지침·도구를 수렴시킨다.

## 파이프라인 (문서별)

```
A1 생성 ──► docNN.hwpx (+manifest)         src/hwpx_gen.py         [크로스플랫폼]
A2 수정 ──► docNN.fixed.hwpx               microalign(AB)+linefix(R-D)  [R-D=로컬 COM]
A3 리뷰 ──► 규칙별 PASS/FAIL               harness/codereview.py   [COM 불필요]
A4 비전 ──► 단절수 0 + 육안 판정           harness/vision.py       [로컬 COM]
A5 자기수정 ► 엔진/도구 패치 or SKILL.md 제안(승인)  [모델/사람]
```

수렴: 전 문서가 (A3 all-PASS) AND (A4 단절 0) = 클린 라운드. **클린 2연속** → 확정.
코드/스펙 변경 시 카운터 0 리셋(회귀 재검증). 상세 = 각 `agents/A*.md`.

## 실행

```bash
# 클라우드/리눅스(점검): A3 코드리뷰만, COM 불필요
python harness/run_harness.py --work tests/_work --manifest tests/_work/manifest.json --review-only

# 로컬(Windows + Hancom Office 2024): 전체 파이프라인(R-D + 비전)
python src/hwpx_gen.py --n 20 --seed 42 --profile messy --out tests/_work/
python harness/run_harness.py --work tests/_work --manifest tests/_work/manifest.json
```

## 로컬 전제

Windows + Hancom Office 2024(COM) + `pip install -r requirements.txt`.
`render_*`/R-D 는 win32com 필요(리눅스 import 는 되나 호출은 COM 부재로 실패 = 의도된 경계).
