# A3 — 코드리뷰 검증 (Reviewer)

**역할**: 원본↔수정 HWPX 의 OWPML diff 정합을 논리/코드 관점에서 검사. 렌더 불필요(COM 무관).
도구 = `harness/codereview.py`.

**입력**: `src.hwpx`, `fixed.hwpx`, (선택) `manifest.json`+`doc`.
**출력**: 규칙별 PASS/FAIL + 결함 리스트(dict). `VERDICT: PASS/FAIL`.

**검증기준**:
- **무결성**: 텍스트 불변 / zip testzip / hwpx open.
- **비파괴**: 원본 charPr in-place 변형 0(추가만) / mimetype ZIP_STORED 선두.
- **R-A**: (BULLET+usage≥2+음수 intent) paraPr → intent 0; 그 외 intent 불변(양수·수동 글머리 함정 불가침).
- **R-B**: 새 bold 선두 run 은 모두 (괄호 토큰 + 본문최빈 스타일 + 잔여≥5) 만족; 제목·잔여<5 함정 미bold.
- **R-D**: 렌더 없이 **비파괴성만**(단절 실측은 A4). spacing 쌍둥이 deepcopy·새 id, 원본 불변.
- **manifest 교차확인**: 기대 대상 반영·함정 불변.

결함은 A5 자기수정 입력.
