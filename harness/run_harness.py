# -*- coding: utf-8 -*-
"""
harness/run_harness.py - 5-에이전트 자기개선 하네스 오케스트레이터 (결정적 백본).

파이프라인(문서별): A2 독립수정 -> A3 코드리뷰 -> A4 비전 -> 집계 -> loop-until-dry(클린 2연속).
A1 생성(hwpx_gen.py)·A5 자기수정(모델/사람)·A4 의미판정(비전 모델)은 이 백본이 호출/공급하는
단계이며, 각 에이전트 역할·입출력은 harness/agents/*.md 스펙을 따른다.

이 백본이 결정적으로 수행하는 것:
  A2 : hwpx_microalign(AB) -> hwpx_linefix(R-D 렌더-루프)  [R-D 는 로컬 COM]
  A3 : harness.codereview.review(원본, 수정, manifest)     [COM 불필요]
  A4 : harness.vision.split_report(수정) 단절 실측 + PNG 렌더 [로컬 COM]
집계 : 문서별 (A3 all-PASS) AND (A4 splits==0) => 클린.
수렴 : 전 문서 클린 라운드가 2연속이면 확정. 결함 발생/코드·스펙 변경 시 카운터 0 리셋.

**로컬 전용**(R-D·비전은 COM). 클라우드에서는 --review-only 로 A3 만 구동(COM 불필요) 가능.

CLI:
  python harness/run_harness.py --work tests/_work --manifest tests/_work/manifest.json
  python harness/run_harness.py --work tests/_work --manifest ... --review-only   # A3만(클라우드 가능)
"""
import sys, os, json, argparse, subprocess, glob, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "src")
sys.path.insert(0, SRC); sys.path.insert(0, HERE)
import codereview  # noqa: E402

PY = sys.executable


def log(m): print(m, flush=True)


def a2_fix(src_hwpx, fixed_hwpx, review_only=False):
    """A2 독립수정: AB(microalign) 후 R-D(linefix, 로컬 COM). review_only 면 AB 만."""
    shutil.copyfile(src_hwpx, fixed_hwpx)
    # R-A/R-B (COM 불필요) — in-place 갱신
    tmp = fixed_hwpx + ".ab"
    subprocess.run([PY, os.path.join(SRC, "hwpx_microalign.py"), fixed_hwpx, tmp, "--rules", "AB"],
                   check=True)
    shutil.move(tmp, fixed_hwpx)
    if review_only:
        return {"rd": "skipped(review-only)"}
    # R-D 렌더-루프 (로컬 COM) — 파일 in-place 수정
    r = subprocess.run([PY, os.path.join(SRC, "hwpx_linefix.py"), fixed_hwpx],
                       capture_output=True, text=True)
    return {"rd_rc": r.returncode, "rd_tail": r.stdout.strip().splitlines()[-1:] if r.stdout else []}


def a3_review(src_hwpx, fixed_hwpx, manifest_entry):
    res = codereview.review(src_hwpx, fixed_hwpx, manifest_entry)
    allok, txt = codereview.summarize(res)
    return allok, res, txt


def a4_vision(fixed_hwpx):
    """A4 단절 실측(로컬 COM). 클라우드/COM 부재 시 예외를 잡아 unknown 처리."""
    try:
        import vision
        rep = vision.split_report(fixed_hwpx)
        return rep["count"] == 0, rep
    except Exception as e:
        return None, {"error": repr(e)}


def run(workdir, manifest_path, review_only=False, max_rounds=8):
    with open(manifest_path, encoding="utf-8") as f:
        manifest = json.load(f)
    docs = [d["file"] for d in manifest.get("docs", [])]
    entry_of = {d["file"]: d for d in manifest.get("docs", [])}

    consecutive_clean = 0
    round_no = 0
    while consecutive_clean < 2 and round_no < max_rounds:
        round_no += 1
        log("\n===== ROUND %d (need 2 consecutive clean) =====" % round_no)
        round_clean = True
        for doc in docs:
            src = os.path.join(workdir, doc)
            if not os.path.exists(src):
                log("[%s] MISSING source, skip" % doc); round_clean = False; continue
            fixed = os.path.join(workdir, doc.replace(".hwpx", ".fixed.hwpx"))
            a2 = a2_fix(src, fixed, review_only=review_only)
            a3_ok, a3_res, a3_txt = a3_review(src, fixed, entry_of.get(doc))
            a4_ok, a4_rep = a4_vision(fixed) if not review_only else (None, {"skipped": True})
            clean = a3_ok and (a4_ok is True)
            if review_only:
                clean = a3_ok  # 비전 생략 시 A3 만으로 판정(수렴 아님, 점검용)
            log("[%s] A2=%s A3=%s A4=%s => %s"
                % (doc, a2, "PASS" if a3_ok else "FAIL",
                   "n/a" if a4_ok is None else ("splits=%d" % a4_rep.get("count", -1)),
                   "CLEAN" if clean else "DIRTY"))
            if not a3_ok: log(a3_txt)
            if not clean: round_clean = False

        if round_clean:
            consecutive_clean += 1
            log("[round %d] ALL CLEAN (%d/2 consecutive)" % (round_no, consecutive_clean))
        else:
            if consecutive_clean:
                log("[round %d] defects -> reset consecutive_clean %d->0" % (round_no, consecutive_clean))
            consecutive_clean = 0
            log(">>> A5 자기수정 필요: 위 FAIL 원인 분석 -> 엔진/도구/프롬프트 패치(자동) 또는 "
                "SKILL.md 변경 제안(사용자 승인). 패치 후 다음 라운드에서 재검증(카운터 리셋 유지).")

    converged = consecutive_clean >= 2
    log("\n===== %s (rounds=%d) =====" % ("CONVERGED (clean x2)" if converged else "NOT CONVERGED", round_no))
    return converged


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--work", default="tests/_work")
    ap.add_argument("--manifest", default=None)
    ap.add_argument("--review-only", action="store_true", help="A3만 구동(COM 불필요, 클라우드 점검용)")
    ap.add_argument("--max-rounds", type=int, default=8)
    args = ap.parse_args()
    manifest = args.manifest or os.path.join(args.work, "manifest.json")
    ok = run(args.work, manifest, review_only=args.review_only, max_rounds=args.max_rounds)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
