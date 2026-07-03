# A1 — 생성 (Generator)

**역할**: 표·캡션·주석 포함 난잡 합성 HWPX 를 생성하고 기대치(manifest)를 낸다. 민감데이터 0.
도구 = `src/hwpx_gen.py`.

**입력**: `seed`, `n`, `profile`(`messy` | `minimal-split`), `out` 디렉토리.
**출력**: `out/docNN.hwpx`(profile=minimal-split & n=1 이면 `testdoc.hwpx`) + `out/manifest.json`.
manifest.docs[i] = `{file, paras, validate_ok, expect:{R-A[], R-B[], R-D_candidates[], traps[]}}`.

**심는 것**: R-A 대상(공유 paraPr, BULLET+음수 intent, ≥2 멤버) / R-B 대상(본문 최빈 스타일 선두 괄호,
잔여≥5) / R-D 후보(긴 무공백 한글 어절). **함정**: 제목(비본문 스타일 괄호)·잔여<5·양수 intent 글머리·수동 글머리.

**검증기준**: 각 문서 `HwpxDocument.open` 가능·`validate().ok`, 심은 R-A/R-B 대상이 규칙 엔진으로 실제
발화(교차확인은 A3), 함정은 미발화. **주의**: R-D 단절은 렌더 정본 — manifest 는 '후보'로만 표기.
A2 에는 manifest 를 **비공개**(정보 격리).
