# HWPX 본문 미세정렬 자기검증·자기개선 하네스 완성 계획

## Context (왜 이 작업인가)

`hwpx-master-harness` 레포의 목표는 hwpx-master가 다음 문서 작업 시 4규칙(R-A/R-B/R-C→R-D)을
**절대 실수 없이 자동 수행**하도록, 규칙별 실현성을 실측·적대검증한 뒤 지침·도구를 수렴·확정하는
self-verifying 하네스를 구축하는 것이다.

**탐색으로 확인된 실제 레포 상태(배경 설명과 중요한 불일치 존재):**

- **이미 존재**: `README.md`(4규칙 스펙의 유일한 현행 근거, L9–23), `src/hwpx_microalign.py`(R-A/R-B
  엔진+R-C flag+charPr 유틸, 자가검증 포함), `src/hwpx_fill_lib.py`(OWPML 조작 헬퍼),
  `src/hwpx_linefix.py`(R-D 렌더-인-루프 엔진, 최신 커밋 `b4bd677`에서 추가·완전 구현).
- **부재(구축 필요)**: `SKILL.md`(README L59가 수렴 목표로만 언급), `.omc/specs`, `tests/`·
  `tests/_work/testdoc.hwpx`, **20문서 합성 생성기**, **5-에이전트 하네스**, 의존성 매니페스트, **git 원격**.
- **엔진 미완 검증**: `hwpx_linefix.py`의 fresh-Dispatch 개정판이 **한 번도 완주 검증 안 됨**(E58 관문 미통과).
- **엔진 코드 갭**: README L28이 약속한 텍스트불변 회귀·`zipfile.testzip()`·`HwpxDocument.open()`
  무결성 검사가 `hwpx_linefix.py`에 **미구현**(microalign에는 있음).

**확정 결정(사용자):**
1. 클라우드는 COM-비의존 코드(생성기·하네스·에이전트 스펙·무결성 검사·requirements·SKILL.md 초안)를
   **실제 작성·커밋**한다. 로컬은 COM 렌더 실행만 담당.
2. 산출물 전달 = **원격 추가 후 push**(사용자가 remote URL 제공 → 브랜치 push → 로컬 pull).
3. 자기수정 게이트: 엔진/도구 코드는 자동 수정, **SKILL.md 변경은 사용자 승인 게이트**.

---

## Phase 0 — 하네스 구축 (CLOUD, 승인 직후 / COM 불필요)

클라우드에서 렌더가 필요 없는 모든 코드를 작성·커밋하고 원격에 push. 로컬 실행의 전제.

### 0.1 의존성·문서 정합
- **`requirements.txt` 생성**: `lxml`, `PyMuPDF`(fitz), `pywin32`(win32com), `hwpx`(python-hwpx) 핀.
  win32com/pywin32는 Windows-only 주석 명시.
- **`README.md` 수정**: L43 stale 문구(`hwpx_linefix.py` "구축 예정") 정정 → "구현 완료, 완주 검증 대기".
  SKILL.md를 규칙 단일 근거로 참조하도록 링크.
- **`SKILL.md` 초안 생성(승인 게이트)**: README L9–23의 4규칙을 정본 스펙으로 이관(R-A 글머리 정렬 /
  R-B 괄호 주제어 bold / R-C→R-D 흡수 / R-D 방향최적 자간 단절방지+글머리 제외+loop-until-clean).
  이후 하네스의 자기수정이 이 파일을 바꿀 때만 사용자 승인 요구. **초안 커밋은 하되, 하네스가 제안하는
  변경 diff는 별도 승인 절차**.

### 0.2 엔진 무결성 보강 — `src/hwpx_linefix.py` 수정 (핵심)
현재 엔진은 비파괴 repack만 있고 README L28의 검증 3종이 없다. `hwpx_microalign.py`의 검증 패턴을 재사용해 추가:
- **텍스트불변 회귀**: `doc_text(sroot)` before/after 동일성(microalign L44 `doc_text`, L260–261 패턴 재사용).
- **ZIP 무결성**: `zipfile.ZipFile(dst).testzip() is None`.
- **HWPX 개방성**: `HwpxDocument.open(dst)` paras>0.
- 매 `apply_span_spacing`+repack 직후(또는 fix_document 종료 시) 이 게이트를 강제, 실패 시 롤백·중단.
- (선택, 병목 완화) 한 iter당 **겹치지 않는 다중 split 동시 조절** + magnitude를 선형(4→20) 대신
  **이분 탐색**으로 전환해 재렌더 횟수 축소. 단, 연쇄 파급 때문에 조절 후 재렌더 실측은 유지.

### 0.3 합성 문서 생성기 — `src/hwpx_gen.py` 생성
- python-hwpx로 표·캡션·주석 포함 **난잡 문서** 생성. `hwpx_fill_lib.py`의 `add_table_row`,
  `set_cell_lines`, `dup_para`, `set_para_text`, `repack`을 빌딩블록으로 재사용.
- **의도적 결함 심기**: (a) 긴 무공백 한글 어절을 텍스트폭 경계 근처에 배치해 **단절 후보** 유도,
  (b) 음수 first-line intent 자동 글머리 그룹(R-A 대상), (c) 본문 선두 괄호 주제어(R-B 대상),
  (d) 배제 케이스(제목/표헤더/캡션/양수 들여쓰기/수동 글머리 — 규칙이 **건드리면 안 되는** 함정).
- **manifest.json 동반 출력**: 문서별 기대치(어느 문단에 R-A/R-B가 발화해야/하지 말아야 하는지,
  단절 후보 위치). **주의**: 단절은 렌더 정본이라 생성 시점엔 "후보"일 뿐 — 확정은 첫 렌더에서(리스크 참조).
- CLI: `python hwpx_gen.py --n 20 --seed S --out tests/_work/`. testdoc용 `--n 1 --profile minimal-split`.
- 산출물 `*.hwpx`는 gitignore 대상 → **생성기 코드만 커밋**, 문서는 로컬 재생성.

### 0.4 5-에이전트 하네스 코드 — `harness/` 생성
- `harness/run_harness.py`: 오케스트레이터(문서 집합 → 5단계 파이프라인 → loop-until-dry).
- `harness/codereview.py`: OWPML diff 정합 검사(COM 불필요, 로직만).
- `harness/vision.py`: COM 렌더→PDF→PNG(fitz `get_pixmap`)→비전 판정 어댑터(로컬 실행).
- `harness/agents/*.md`: 5개 에이전트 프롬프트 스펙(아래 §에이전트 명세). 입출력 스키마·검증기준 포함.
- 커밋 후 **원격 추가(`git remote add origin <URL>`) → `git push -u origin <branch>`** (사용자 URL 제공).

**Phase 0 완료 기준**: 위 파일 커밋·push 완료, 로컬에서 `pip install -r requirements.txt` +
`python src/hwpx_gen.py` 실행 가능(코드 레벨). SKILL.md 초안은 커밋되되 이후 변경은 승인 게이트.

---

## Phase 1 — 엔진 검증 = E58 관문 (LOCAL, 첫 로컬 마일스톤 / go-no-go)

**이것이 사용자의 첫 로컬 실행이자 전체 진행의 관문.**

1. 로컬(Windows + Hancom Office 2024)에서 pull → `pip install -r requirements.txt`.
2. `python src/hwpx_gen.py --n 1 --profile minimal-split --out tests/_work/` → `tests/_work/testdoc.hwpx`
   (단절 어절 1–2개를 담은 최소 문서).
3. `python src/hwpx_linefix.py tests/_work/testdoc.hwpx` 실행.
4. **관문 통과 기준**:
   - 첫 COM `Open`이 E58(RPC 서버 오염) 없이 성공하거나, 발생 시 fresh-Dispatch 재시도로 자동 복구.
   - `fix_document`가 **1회 완주** → `{"clean": True, ...}` 반환(단절 0까지 loop).
   - 0.2에서 추가한 무결성 3종 게이트 PASS(텍스트불변·testzip·hwpx open).
   - 렌더 PDF의 해당 어절이 실제로 1줄화(fitz `detect_splits`가 0).
5. **실패 시 대응**: E58 지속 → §리스크 COM 오염 대응 절차. 매핑 실패(`unmappable`) → `map_split`
   window/키 로직 보정(자기수정 대상). 이 관문을 통과하기 전엔 파일럿 진입 금지.

**산출**: 완주 로그 + before/after 렌더 PNG 1쌍(육안·비전 확인).

---

## Phase 2 — 파일럿 (LOCAL, 3–5문서)

1. `python src/hwpx_gen.py --n 5 --seed <fixed> --out tests/_work/` → doc01–05 + manifest.
2. 3–5문서에 **5-에이전트 파이프라인 1라운드** 수동/오케스트레이터로 구동(§에이전트 명세).
3. **목표**: 파이프라인 배선·에이전트 입출력·검증기준을 소규모에서 실증. 각 규칙이 함정 케이스를
   건드리지 않는지(R-A 양수 intent/수동 글머리 불가침, R-B 제목/캡션 배제) 코드리뷰+비전 이중검증.
4. 여기서 발견되는 결함은 자기수정(§자기개선)으로 엔진/도구/프롬프트에 반영. 병목·COM 안정성 실측.

**진입 게이트**: Phase 1 관문 통과 필수. **완료 기준**: 파일럿 전 문서가 코드리뷰+비전 **1라운드 클린**,
파이프라인이 오케스트레이터로 무인 구동 가능.

---

## Phase 3 — 풀가동 (LOCAL, 20문서, loop-until-dry)

1. `python src/hwpx_gen.py --n 20 --seed <fixed> --out tests/_work/` → doc01–20 + manifest.
2. `harness/run_harness.py`로 20문서 **loop-until-dry** 구동.
3. **수렴 조건**: 아래 §수렴 정의의 "클린 2연속" 달성.
4. **산출**: 수렴 로그, 문서별 before/after 렌더 PNG, 확정된 SKILL.md(승인 반영)·도구 버전, 최종 커밋·push.

**완료 기준(README 수용기준 충족)**: 20문서 전부 단절 0 + R-A/R-B 정확 + 텍스트불변·무결성, 코드리뷰+비전
이중검증을 **연속 2라운드** 통과, 지침·도구 수렴 확정.

---

## 5-에이전트 명세 (역할·입출력·검증기준)

각 에이전트는 분리(정보 격리)된다. 특히 **생성↔독립수정은 manifest 비공유**(수정은 결과 HWPX만 본다).

### A1. 생성 (Generator) — CLOUD 코드 / LOCAL 실행
- **입력**: seed, n, profile. **출력**: `docNN.hwpx` + `manifest.json`(기대 발화/배제/단절 후보).
- **검증기준**: 생성물이 `HwpxDocument.open` 가능·zip 무결, 의도한 결함 패턴(단절 후보·음수 intent
  글머리·괄호 주제어·함정 케이스)을 텍스트/paraPr 스캔으로 포함 확인. **단, 단절 확정은 렌더 후**.

### A2. 독립 수정 (Fixer) — LOCAL(R-D COM)
- **입력**: `docNN.hwpx`(manifest 비공개). **출력**: `docNN.fixed.hwpx` + change-log(규칙별 적용 내역).
- **처리**: R-A/R-B = `hwpx_microalign.py --rules AB`; R-D = `hwpx_linefix.py`(렌더-루프).
- **검증기준**: 텍스트불변·testzip·hwpx open PASS, **멱등성**(재실행 시 추가 변경 0).

### A3. 코드리뷰 검증 (Reviewer) — COM 불필요(로컬 아티팩트 대상)
- **입력**: 원본·수정 HWPX의 OWPML diff + change-log. **출력**: 규칙별 PASS/FAIL + 결함 리스트.
- **검증기준(정합성)**: R-A는 음수 intent 자동 글머리 paraPr만(≥2 멤버) 0으로; R-B는 선두 괄호
  토큰만, 본문최빈 스타일+잔여≥5 게이트 준수, bold 쌍둥이 charPr는 크기·글꼴 동일·새 id; R-D 스팬은
  줄 첫글자~조절 어절, **글머리 제외**, spacing 쌍둥이 deepcopy·새 id; mimetype ZIP_STORED 선두;
  **텍스트 불변**. 함정 케이스(제목/캡션/양수 intent) 미변경 확인.

### A4. 비전 검증 (Vision) — LOCAL(COM 렌더 정본)
- **입력**: `docNN.fixed.hwpx`. **처리**: COM→PDF→PNG(fitz). **출력**: 페이지별 판정 + 실패 주석.
- **검증기준(렌더 실측)**: 단절 어절 0(육안+`detect_splits`), 자동 글머리 접힌 줄이 첫줄 본문에 정렬,
  괄호 주제어가 진하게, 배제 대상 미변경. **이것이 최종 정본**(linesegarray/manifest보다 우선).

### A5. 자기수정 (Self-update) — CLOUD 코드 수정 / 승인 게이트
- **입력**: A3·A4 결함 리포트. **출력**: 근본원인 분석 + (엔진/도구/프롬프트 자동 패치) 또는
  (SKILL.md 변경 **제안 diff → 사용자 승인 대기**).
- **검증기준**: 패치 후 해당 결함이 재현 안 됨, 회귀 없음(다른 문서 미퇴행). SKILL.md는 승인 전 미적용.

---

## loop-until-dry 수렴 정의 (클린 2연속)

- **1라운드** = 현재 문서 집합 전체를 A2→A3→A4로 1회 통과.
- **클린 라운드** = 집합의 **모든 문서**가 A3(코드리뷰) **및** A4(비전)에서 결함 0 + 단절 0 + 무결성 PASS.
- **수렴** = **클린 라운드 2회 연속**. `consecutive_clean` 카운터: 클린이면 +1, 결함 발생 또는
  **엔진/도구/프롬프트 변경(A5 패치) 발생 시 0으로 리셋**(변경이 회귀를 유발할 수 있으므로 재검증 강제).
- SKILL.md 승인 변경도 스펙 변화이므로 리셋 대상. 카운터 2 도달 → 지침·도구 확정, 최종 커밋.
- **상한/안전장치**: 라운드 상한(예 8)·문서별 `fix_document` `max_global` 상한으로 무한루프 방지.
  상한 도달 시 미해소 문서를 `failed`로 격리·리포트(무언의 truncation 금지).

---

## 리스크 & 대응

### R1. COM 서버 상태 오염 (E58) — 최우선
- **증상**: HwpObject 재사용·`XHwpDocuments` 접근·루프 중 Hwp.exe taskkill → "RPC 서버 사용 불가"
  후 `AttributeError('Open')`.
- **대응**: (현행 준수) 렌더마다 **fresh `Dispatch`+`Open`+`SaveAs`+`Quit`**, `XHwpDocuments` 미접근,
  **taskkill 금지**. 추가: `render_pdf`에 E58 감지 시 `CoUninitialize`→`CoInitialize`→fresh Dispatch로
  **1회 백오프 재시도**; 지속 시 루프 **일시중지 + 사용자에 표면화**(강제종료 대신). Phase 1이 이 관문.

### R2. COM 직렬·상태ful 병목
- **증상**: 렌더가 순차·수 초/회. 20문서 × loop iter × magnitude 시도 = 대량 재렌더 → 벽시계 폭증.
- **대응**: (a) iter당 겹치지 않는 다중 split 동시 조절로 iter 수 감소, (b) magnitude 이분 탐색,
  (c) 변경 없으면 재렌더 skip·PDF 캐시, (d) 백그라운드 순차 큐 + 진행 로그, (e) 파일럿에서 실측 후
  `max_global`·magnitude 상한 튜닝. 병렬 COM은 **금지**(직렬 유지).

### R3. 단절 합성 불확실성(렌더 정본)
- **증상**: 생성기는 linesegarray 신뢰 금지 원칙상 단절을 **보장 못 함**(렌더에서만 확정).
- **대응**: 생성 후 **첫 렌더로 단절 실재 확인**, 없으면 텍스트폭 경계 조정(어절 길이·여백)으로
  재생성. manifest는 "후보"로 표기, A4 비전이 실제값 확정.

### R4. 비파괴 회귀(텍스트·무결성)
- **대응**: 0.2에서 엔진에 텍스트불변·testzip·hwpx open 게이트 추가(현재 미구현). 매 편집 후 강제,
  실패 시 롤백. A2/A3가 이중 확인.

### R5. 매핑 실패(`unmappable`)
- **증상**: `map_split` 키 window(14/10/7/5)가 문단 유일 매칭 실패(반복 텍스트·표 셀).
- **대응**: window 확장·문단 인덱스 힌트 추가를 자기수정(A5) 대상으로. 실패 split은 격리·리포트.

### R6. 산출물 전달(원격 부재)
- **대응**: Phase 0 말미 `git remote add origin <사용자 URL>` → `git push -u origin <branch>`
  (실패 시 지수 백오프 4회 재시도). 로컬은 `git pull`. 원격 미제공 시 대체로 git bundle 산출.

---

## 진행 상태 (2026-07-03 갱신 — Phase 1 통과)

- **Phase 1 완료 (LOCAL, E58 관문 통과)**: 실물 문서(전기기초실습 4.4MB, 1058 시각줄)에서
  렌더-루프 36 iter 완주(189s), **clean=True · unique_fixed=107 어절 · 무결성 3종 PASS · E58 0회**(~96 렌더).
  before 137 -> after 62 단절(잔여 62 = 표 셀/수식 거짓양성 60 + 진동 격리 2 — failed 로 정직 보고).
  before/after 렌더 PNG 육안 대조로 "등의·사용법을·저항의·임피던스의" 등 1줄화 확인.
- **Phase 1 중 발견·자가수정한 결함 5건** (자기수정 게이트: 엔진/도구 = 자동):
  1. `hwpx_gen.py` `ensure_run_style(size=)` — 로컬 hwpx 2.9 미지원 -> `ensure_char_property`(height twin) 로 교체.
  2. `hwpx_gen.py` `set_list_format`/`set_paragraph_format` — 2.9 부재 -> 실물 paraPr 구조(heading BULLET +
     hp:switch case/default margin intent) 직접 주입(`_ensure_bullet_parapr`). microalign R-A 발화+함정 불가침 검증.
  3. python-hwpx `header.element` 는 type hint(ET)와 달리 **lxml** — 생성기 임포트 정합.
  4. `hwpx_linefix.py` unmappable 2회 후 전체 포기 -> **skip-set 격리 + 문단당 1건 배치 + iter 당 렌더 1회**
     (표 셀 거짓양성 61건이 진행을 막던 결함 해소, R2 병목완화 동시 반영).
  5. 같은 문단 widen<->narrow 밀당 **진동** -> 사다리 기억(재발 시 이어감) + 재발 2회 narrow 고정 +
     4회 격리. max_iter 발산이 clean 수렴으로 전환.
- **알려진 잔여(비차단, A5 백로그)**: (a) 수식 개체 인접 단절 4건 unmappable(PDF 텍스트!=OWPML 텍스트),
  (b) 진동 미해소 2어절(장비의·멀티미터·전원 — 문단 재배치 전략 필요), (c) 연속 renders 무간격 시
  간헐 COM Open 실패(엔진 루프는 자연 간격으로 무영향; R1 백오프 재시도 추가 후보),
  (d) minimal-split 합성문서가 비정상 협폭 렌더(생성기 페이지 셋업 부재 — messy 프로파일/실물로 대체).
- **다음 = Phase 2 파일럿**: messy 3~5문서(생성기 검증 완료: R-A 3·R-B 1·R-D 후보 6·함정 6,
  microalign VERDICT PASS) -> 5-에이전트 파이프라인 1라운드.

---

## 진행 상태 (2026-07-02 갱신)

- **Phase 0 완료·커밋됨**: 브랜치 `feat/harness-scaffold`, 커밋 `b0df6c6`
  (SKILL.md·requirements.txt·hwpx_gen.py·harness/ 4파일+agents 5스펙·linefix 무결성 게이트·README 정정).
  클라우드 검증: 생성기 valid 산출, microalign 발화·함정 불가침·멱등, codereview 양성/음성, run_harness
  --review-only 클린 2연속 수렴 — 전부 PASS.
- **다음 단계(즉시 실행 대상)**: 사용자 제공 원격 `https://github.com/pyu4277/HWP_Master.git`에
  1. `git remote add origin https://github.com/pyu4277/HWP_Master.git`
  2. `git push -u origin main` (base 브랜치 공유)
  3. `git push -u origin feat/harness-scaffold` (실패 시 지수 백오프 2s/4s/8s/16s 재시도)
  4. draft PR 생성 (`feat/harness-scaffold` → `main`)
  5. 인증 실패 시 대체: `git bundle create` 로 전달.
- **그 후**: 사용자 로컬 pull → Phase 1 E58 관문(testdoc 렌더-루프 완주) 실행.

---

## 수정/생성 대상 핵심 파일

- **수정** `src/hwpx_linefix.py` — 무결성 3종 게이트 추가(§0.2), (선택) 다중-split·이분탐색 병목완화.
  재사용: `hwpx_microalign.py`의 `doc_text`(L44), 검증 패턴(L259–273), `repack`(L219–228).
- **생성** `src/hwpx_gen.py` — 난잡 합성 생성기(+manifest). 재사용: `hwpx_fill_lib.py` 표/셀/문단 헬퍼.
- **생성** `harness/run_harness.py`, `harness/codereview.py`, `harness/vision.py`, `harness/agents/*.md`.
- **생성** `SKILL.md`(승인 게이트), `requirements.txt`; **수정** `README.md`(L43 stale 정정).
- **로컬 산출(gitignore)** `tests/_work/testdoc.hwpx`, `docNN.hwpx`, 렌더 PDF/PNG.

---

## 검증(엔드투엔드)

1. **엔진 관문(Phase 1)**: `python src/hwpx_linefix.py tests/_work/testdoc.hwpx` → `{"clean": True}`
   + 무결성 3종 PASS + E58 무발생/자동복구. before/after 렌더 PNG로 단절 소멸 확인.
2. **규칙 정합(Phase 2)**: `python src/hwpx_microalign.py <src> <dst> --rules AB --flag-c` VERDICT: PASS,
   함정 케이스 미변경을 A3 코드리뷰로 확인.
3. **파일럿 수렴**: 3–5문서 1라운드 코드리뷰+비전 클린.
4. **풀 수렴(Phase 3)**: `harness/run_harness.py` → 20문서 클린 2연속, 최종 커밋·push.
