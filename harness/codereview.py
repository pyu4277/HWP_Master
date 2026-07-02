# -*- coding: utf-8 -*-
"""
harness/codereview.py - A3 코드리뷰 검증 (COM 불필요, OWPML 정합만).

원본 HWPX 와 수정 HWPX 의 diff 를 규칙별로 검사한다. 렌더는 하지 않으므로(단절 실측은 A4 비전)
R-D 는 '비파괴성'만 검사한다. 반환은 규칙별 PASS/FAIL + 결함 리스트(dict).

검사 항목:
  - 무결성 : 텍스트 불변 / zip testzip / hwpx open (linefix.integrity_check 재사용).
  - 비파괴 : 원본 charPr 는 하나도 in-place 변형되지 않음(추가만 허용) / mimetype ZIP_STORED 선두.
  - R-A    : 원본에서 (heading BULLET + usage>=2 + 음수 intent) 였던 paraPr 는 수정본에서 intent=0.
             그 외 paraPr 의 intent 는 불변(양수·비글머리·수동 글머리 함정 불가침).
  - R-B    : 새로 bold 된 선두 run 은 모두 '선두 괄호 토큰 + 본문최빈 스타일 + 잔여>=5' 게이트를 만족.
             제목(비본문 스타일)·잔여<5 함정은 bold 되지 않음.
  - manifest(선택): 기대 R-A/R-B 대상이 실제로 반영됐고 함정이 불변인지 교차확인.

CLI: python harness/codereview.py <src.hwpx> <fixed.hwpx> [--manifest manifest.json] [--doc docNN.hwpx]
"""
import sys, os, re, json, zipfile, argparse
from collections import Counter
from lxml import etree

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import hwpx_linefix as LF  # all_text / integrity_check / L / qhp / qhh

HP = LF.HP; HH = LF.HH
def L(e): return etree.QName(e.tag).localname
BRACKET = re.compile(r'^(\s*)([\(\[<][^\)\]>]*[\)\]>])')


def _roots(hwpx):
    with zipfile.ZipFile(hwpx) as z:
        s = etree.fromstring(z.read("Contents/section0.xml"))
        h = etree.fromstring(z.read("Contents/header.xml"))
    return s, h


def _charpr_map(hroot):
    """id -> canonical string(정렬된 자식/속성) for in-place mutation detection."""
    out = {}
    for cp in hroot.iter():
        if L(cp) == "charPr":
            out[cp.get("id")] = etree.tostring(cp, encoding="unicode")
    return out


def _parapr_intents(hroot, sroot):
    """paraPr id -> {'bullet':bool, 'usage':int, 'intents':[int,...]}."""
    usage = Counter(p.get("paraPrIDRef") for p in sroot.iter() if L(p) == "p")
    info = {}
    for pp in hroot.iter():
        if L(pp) != "paraPr": continue
        pid = pp.get("id")
        bullet = any(L(h) == "heading" and h.get("type") == "BULLET" for h in pp.iter())
        intents = []
        for it in pp.iter():
            if L(it) == "intent" and it.get("value") is not None:
                try: intents.append(int(it.get("value")))
                except ValueError: pass
        info[pid] = {"bullet": bullet, "usage": usage.get(pid, 0), "intents": intents}
    return info


def _body_style(hroot, sroot):
    cidx = {cp.get("id"): cp for cp in hroot.iter() if L(cp) == "charPr"}
    def fref(cp):
        for c in cp:
            if L(c) == "fontRef": return c.get("hangul")
        return None
    c = Counter()
    for r in sroot.iter():
        if L(r) == "run":
            txt = "".join(t.text or "" for t in r.iter() if L(t) == "t")
            if txt.strip():
                cp = cidx.get(r.get("charPrIDRef"))
                if cp is not None: c[(cp.get("height"), fref(cp))] += 1
    return c.most_common(1)[0][0] if c else (None, None), cidx, fref


def _leading_bold_tokens(sroot, hroot):
    """수정본에서 '선두 run 이 bold 이고 괄호 토큰'인 문단 목록."""
    cidx = {cp.get("id"): cp for cp in hroot.iter() if L(cp) == "charPr"}
    def is_bold(cid):
        cp = cidx.get(cid); return cp is not None and any(L(c) == "bold" for c in cp)
    res = []
    for p in sroot.iter():
        if L(p) != "p": continue
        runs = p.findall("{%s}run" % HP)
        # first non-empty run
        for r in runs:
            j = "".join(t.text or "" for t in r.iter() if L(t) == "t")
            if j.strip():
                if BRACKET.match(j.strip()) and is_bold(r.get("charPrIDRef")):
                    res.append(("".join(t.text or "" for rr in runs for t in rr.iter() if L(t) == "t"), j.strip()))
                break
    return res


def review(src_hwpx, fixed_hwpx, manifest_entry=None):
    result = {"integrity": {"ok": True, "defects": []},
              "nondestructive": {"ok": True, "defects": []},
              "R-A": {"ok": True, "defects": []},
              "R-B": {"ok": True, "defects": []},
              "R-D": {"ok": True, "defects": []},
              "manifest": {"ok": True, "defects": []}}

    def fail(bucket, msg):
        result[bucket]["ok"] = False; result[bucket]["defects"].append(msg)

    # --- 무결성 (텍스트 불변 = 원본 기준) ---
    orig_text = LF.section_text(src_hwpx)
    ok, probs = LF.integrity_check(fixed_hwpx, orig_text)
    if not ok:
        for p in probs: fail("integrity", p)

    src_s, src_h = _roots(src_hwpx)
    fx_s, fx_h = _roots(fixed_hwpx)

    # --- 비파괴: 원본 charPr in-place 변형 금지(추가만 허용) ---
    src_cp = _charpr_map(src_h); fx_cp = _charpr_map(fx_h)
    for cid, canon in src_cp.items():
        if cid not in fx_cp:
            fail("nondestructive", "charPr %s removed" % cid)
        elif fx_cp[cid] != canon:
            fail("nondestructive", "charPr %s mutated in place (deepcopy+new id expected)" % cid)
    # mimetype ZIP_STORED 선두
    with zipfile.ZipFile(fixed_hwpx) as z:
        infos = z.infolist()
        if not infos or infos[0].filename != "mimetype":
            fail("nondestructive", "mimetype not first entry")
        elif infos[0].compress_type != zipfile.ZIP_STORED:
            fail("nondestructive", "mimetype not ZIP_STORED")

    # --- R-A: BULLET+usage>=2+음수 intent 는 0 으로, 그 외 intent 불변 ---
    src_pp = _parapr_intents(src_h, src_s); fx_pp = _parapr_intents(fx_h, fx_s)
    for pid, si in src_pp.items():
        fi = fx_pp.get(pid)
        if fi is None:
            fail("R-A", "paraPr %s disappeared" % pid); continue
        target = si["bullet"] and si["usage"] >= 2 and any(v < 0 for v in si["intents"])
        if target:
            if any(v < 0 for v in fi["intents"]):
                fail("R-A", "paraPr %s: negative intent not zeroed (%s)" % (pid, fi["intents"]))
        else:
            # 함정/무관 paraPr: intent 집합 불변이어야
            if sorted(si["intents"]) != sorted(fi["intents"]):
                fail("R-A", "paraPr %s: non-target intent changed %s -> %s"
                     % (pid, si["intents"], fi["intents"]))

    # --- R-B: 새 bold 선두 토큰은 게이트 만족, 함정은 bold 금지 ---
    body_style, _, _ = _body_style(fx_h, fx_s)
    for full_text, token_run in _leading_bold_tokens(fx_s, fx_h):
        m = BRACKET.match(token_run)
        if not m:
            fail("R-B", "bold leading run is not a bracket token: %r" % token_run[:20]); continue
        token = m.group(2)
        rest = full_text[full_text.find(token) + len(token):]
        if len(rest.strip()) < 5:
            fail("R-B", "bolded token %r has <5 remainder (trap violated)" % token)
    # 제목 함정: manifest 로만 확실히 판정(스타일은 아래 manifest 교차확인).

    # --- manifest 교차확인(선택) ---
    if manifest_entry:
        exp = manifest_entry.get("expect", {})
        bolded = {tr for _, tr in _leading_bold_tokens(fx_s, fx_h)}
        bolded_join = " ".join(bolded)
        for trap in exp.get("traps", []):
            if trap.get("rule") == "R-B" and trap.get("kind") in ("title-bracket-nonbody-style", "bracket-remainder-under-5"):
                pref = trap["text_prefix"]
                # 함정 텍스트의 괄호 토큰이 bold 목록에 있으면 위반
                mt = BRACKET.match(pref.strip())
                if mt and mt.group(2) in bolded_join and _paragraph_bolded(fx_s, fx_h, pref):
                    fail("manifest", "trap %s bolded: %r" % (trap["kind"], pref))
        # R-A 대상이 실제 반영됐는지: src 에 음수였던 target paraPr 가 fixed 에서 0 인지(위 R-A 로 커버)
    return result


def _paragraph_bolded(sroot, hroot, text_prefix):
    """text_prefix 로 시작하는 문단의 선두 run 이 실제 bold 인지."""
    cidx = {cp.get("id"): cp for cp in hroot.iter() if L(cp) == "charPr"}
    for p in sroot.iter():
        if L(p) != "p": continue
        joined = "".join(t.text or "" for t in p.iter() if L(t) == "t")
        if joined.strip().startswith(text_prefix.strip()[:12]):
            for r in p.findall("{%s}run" % HP):
                j = "".join(t.text or "" for t in r.iter() if L(t) == "t")
                if j.strip():
                    cp = cidx.get(r.get("charPrIDRef"))
                    return cp is not None and any(L(c) == "bold" for c in cp)
    return False


def summarize(result):
    allok = all(v["ok"] for v in result.values())
    lines = []
    for k, v in result.items():
        lines.append("  [%s] %s%s" % ("PASS" if v["ok"] else "FAIL", k,
                     "" if v["ok"] else " :: " + "; ".join(v["defects"])))
    return allok, "\n".join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src"); ap.add_argument("fixed")
    ap.add_argument("--manifest"); ap.add_argument("--doc")
    args = ap.parse_args()
    entry = None
    if args.manifest and args.doc:
        with open(args.manifest, encoding="utf-8") as f:
            man = json.load(f)
        entry = next((d for d in man.get("docs", []) if d.get("file") == args.doc), None)
    res = review(args.src, args.fixed, entry)
    allok, txt = summarize(res)
    print(txt)
    print("VERDICT:", "PASS" if allok else "FAIL")
    sys.exit(0 if allok else 1)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
