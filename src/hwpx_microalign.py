# -*- coding: utf-8 -*-
"""
hwpx_microalign.py - HWPX 본문 미세정렬 3규칙 (검증 후 승격, 2026-07-02)

R-A 글머리 내어쓰기 정렬  : 자동 글머리(heading BULLET) paraPr 의 음수 intent 만 0 으로(양수/수동/제목 불가침). FULL 자동.
R-B 괄호 주제어 bold      : 본문 선두 (),[],<> 주제어 토큰만 bold(본문최빈스타일 게이트 + 잔여>=5 게이트). FULL 자동.
R-C 자간 2줄->1줄          : 후보 플래깅만(자동 붕괴 금지). 실제 붕괴 결정은 한글 육안 확정(렌더러 부재 E36).

비파괴: 공유 charPr 는 deepcopy 후 새 id. mimetype ZIP_STORED 선두 repack. 적용 후 무결성+텍스트불변 자가검증.
근거: workflow wf_58a994b8 (feasibility+adversarial verify), ground truth = 사용자 한/글 완성본.
"""
import sys, os, re, copy, zipfile, argparse
from collections import Counter
from lxml import etree

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HH = "http://www.hancom.co.kr/hwpml/2011/head"
def L(e): return etree.QName(e.tag).localname
def qhp(t): return "{%s}%s" % (HP, t)
def qhh(t): return "{%s}%s" % (HH, t)

BRACKET = re.compile(r'^(\s*)([\(\[<][^\)\]>]*[\)\]>])')

def single_t_run(src_run, charpr_id, text):
    """Deepcopy a run, keep ONLY its first <t> (drop all other children incl. control siblings
    like lineBreak/tab/ctrl so they are not duplicated), set charPrIDRef + text. Returns new run."""
    r = copy.deepcopy(src_run)
    if charpr_id is not None:
        r.set("charPrIDRef", str(charpr_id))
    kept = None
    for c in list(r):
        if L(c) == "t" and kept is None:
            kept = c
        else:
            r.remove(c)
    if kept is None:
        kept = etree.SubElement(r, qhp("t"))
    kept.text = text
    return r

def load(path, name):
    with zipfile.ZipFile(path) as z: return z.read(name)
def ptext_p(p): return "".join(t.text or "" for t in p.iter() if L(t) == "t")
def doc_text(sroot): return "".join(ptext_p(p) for p in sroot.iter() if L(p) == "p")

# ---------------- charPr helpers ----------------
def charpr_index(hroot):
    return {cp.get("id"): cp for cp in hroot.iter() if L(cp) == "charPr"}
def has_bold(cp): return any(L(c) == "bold" for c in cp)
def font_ref(cp):
    for c in cp:
        if L(c) == "fontRef": return c.get("hangul")
    return None
def charpr_container(hroot):
    for cp in hroot.iter():
        if L(cp) == "charPr":
            return cp.getparent()
    return None
def sig_no_bold(cp):
    """signature of charPr ignoring bold: height + ordered non-bold children (name+attrs)."""
    parts = [("@height", cp.get("height"))]
    for c in cp:
        if L(c) == "bold": continue
        parts.append((L(c), tuple(sorted(c.attrib.items()))))
    return tuple(parts)

def get_bold_twin(hroot, base_id, cidx):
    """Return id of a charPr identical to base_id but with <hh:bold/>. Reuse if exists, else generate."""
    base = cidx.get(base_id)
    if base is None: return None
    if has_bold(base): return base_id
    target = sig_no_bold(base)
    for cid, cp in cidx.items():
        if has_bold(cp) and sig_no_bold(cp) == target:
            return cid
    # generate new twin: deepcopy base, insert <hh:bold/> right after 'offset' (schema position)
    newcp = copy.deepcopy(base)
    ids = [int(c.get("id")) for c in cidx.values() if c.get("id") and c.get("id").isdigit()]
    newid = str(max(ids) + 1)
    newcp.set("id", newid)
    # find insertion index: after last of [fontRef,ratio,spacing,relSz,offset]
    order = ["fontRef", "ratio", "spacing", "relSz", "offset"]
    kids = list(newcp)
    idx = 0
    for i, c in enumerate(kids):
        if L(c) in order: idx = i + 1
    bold_el = etree.Element(qhh("bold"))
    newcp.insert(idx, bold_el)
    cont = charpr_container(hroot)
    cont.append(newcp)
    cnt = cont.get("itemCnt")
    if cnt is not None: cont.set("itemCnt", str(int(cnt) + 1))
    cidx[newid] = newcp
    return newid

def body_mode_style(hroot, sroot, cidx):
    """(height, fontRef) most common among non-empty section runs."""
    c = Counter()
    for r in sroot.iter():
        if L(r) == "run":
            txt = "".join(t.text or "" for t in r.iter() if L(t) == "t")
            if txt.strip():
                cp = cidx.get(r.get("charPrIDRef"))
                if cp is not None:
                    c[(cp.get("height"), font_ref(cp))] += 1
    return c.most_common(1)[0][0] if c else (None, None)

# ---------------- Rule A: auto-bullet negative-intent normalization ----------------
def rule_a(hroot, sroot, report):
    usage = Counter(p.get("paraPrIDRef") for p in sroot.iter() if L(p) == "p")
    changed = 0
    for pp in hroot.iter():
        if L(pp) != "paraPr": continue
        pid = pp.get("id")
        is_bullet = any(L(h) == "heading" and h.get("type") == "BULLET" for h in pp.iter())
        if not is_bullet: continue
        if usage.get(pid, 0) < 2: continue            # only real list groups (>=2 members)
        for blk in pp.iter():
            if L(blk) in ("case", "default"):
                for m in blk.iter():
                    if L(m) == "margin":
                        it = m.find("./{*}intent")
                        if it is not None and it.get("value") and int(it.get("value")) < 0:
                            report.append("  R-A paraPr %s %s intent %s -> 0" % (pid, L(blk), it.get("value")))
                            it.set("value", "0"); changed += 1
    return changed

# ---------------- Rule B: leading bracket 주제어 bold ----------------
def is_title_or_excluded(p, run0_cp, body_style):
    """Exclude if leading run style != body mode (title/header/subheading/foreign font)."""
    if run0_cp is None: return True
    return (run0_cp.get("height"), font_ref(run0_cp)) != body_style

def rule_b(hroot, sroot, cidx, body_style, report):
    applied = 0
    for p in sroot.iter():
        if L(p) != "p": continue
        runs = p.findall(qhp("run"))
        # first run with non-empty t
        r0 = None; t0 = None
        for r in runs:
            ts = r.findall(qhp("t"))
            joined = "".join(t.text or "" for t in ts)
            if joined.strip():
                r0 = r; t0 = ts[0] if ts else None; break
        if r0 is None or t0 is None or t0.text is None: continue
        text = t0.text
        m = BRACKET.match(text)
        if not m: continue
        # gate (a): body-mode style
        r0_cp = cidx.get(r0.get("charPrIDRef"))
        if is_title_or_excluded(p, r0_cp, body_style): continue
        if r0_cp is not None and has_bold(r0_cp): continue    # idempotence: never re-bold an already-bold lead
        if any(L(c) != "t" for c in r0): continue             # safety: only split runs whose children are all <t>
        lead_ws, token = m.group(1), m.group(2)
        rest_in_run = text[len(m.group(0)):]
        # gate (b): remainder body text >= 5 chars (whole paragraph after token)
        full_rest = rest_in_run + "".join(
            "".join(x.text or "" for x in rr.iter() if L(x) == "t")
            for rr in runs[runs.index(r0)+1:])
        if len(full_rest.strip()) < 5: continue
        # resolve/generate bold twin of the leading run's charPr
        bold_id = get_bold_twin(hroot, r0.get("charPrIDRef"), cidx)
        if bold_id is None: continue
        # split: set original run t -> rest_in_run; create bold run(token); optional ws run
        t0.text = rest_in_run
        bold_run = single_t_run(r0, bold_id, token)
        r0.addprevious(bold_run)
        if lead_ws:
            ws_run = single_t_run(r0, r0.get("charPrIDRef"), lead_ws)  # body style, only the leading ws
            bold_run.addprevious(ws_run)
        report.append("  R-B bold '%s' | rest '%s...'" % (token, full_rest.strip()[:30]))
        applied += 1
    return applied

# ---------------- Rule C: 2-line candidate flagging (no auto-apply) ----------------
def rule_c_flag(sroot, cidx, body_style, report):
    """Flag ONLY conservative candidates: exactly-2-line prose whose 2nd line is a small tail
    (tail_ratio = line1_chars / line0_chars <= 0.35). Never auto-applies; human 한글 confirm required.
    Rationale: small tail => few chars to absorb => most likely to collapse within 자간 -20; large
    tail (near-full 2nd line) needs more than -20 (verified unreliable). Table headers / short labels
    excluded by min length. Bold-leading (R-B) paragraphs excluded (cannot tighten bold token)."""
    cands = 0
    text = None
    for p in sroot.iter():
        if L(p) != "p": continue
        lsa = p.find(qhp("linesegarray"))
        if lsa is None: continue
        segs = lsa.findall(qhp("lineseg"))
        if len(segs) != 2: continue                     # exactly 2 lines
        text = ptext_p(p)
        txt = text.strip()
        if len(txt) < 25: continue                       # prose only (skip labels/headers/short cells)
        if any(L(e) == "lineBreak" for e in p.iter()): continue   # skip forced break
        try:
            n0 = int(segs[1].get("textpos"))              # chars on line 0
        except (TypeError, ValueError):
            continue
        n1 = len(text) - n0
        if n0 <= 0: continue
        tail = n1 / n0
        if tail > 0.35: continue                          # only small-tail (barely overflows)
        # skip bold-leading (Rule B territory)
        runs = p.findall(qhp("run"))
        lead_bold = False
        for r in runs:
            j = "".join(t.text or "" for t in r.iter() if L(t) == "t")
            if j.strip():
                cp = cidx.get(r.get("charPrIDRef"))
                if cp is not None and has_bold(cp) and BRACKET.match(j):
                    lead_bold = True
                break
        if lead_bold: continue
        cands += 1
        report.append("  R-C candidate(tail=%.2f, 한글확인필요): %s" % (tail, txt[:52]))
    return cands

# ---------------- repack + self-verify ----------------
def repack(src, dst, datamap):
    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        if "mimetype" in datamap:
            zi = zipfile.ZipInfo("mimetype"); zi.compress_type = zipfile.ZIP_STORED
            zout.writestr(zi, datamap["mimetype"])
        for n in names:
            if n == "mimetype": continue
            zout.writestr(n, datamap[n])

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src"); ap.add_argument("dst")
    ap.add_argument("--rules", default="AB", help="subset of A,B to APPLY (C is always flag-only)")
    ap.add_argument("--flag-c", action="store_true", help="also report R-C 2-line candidates")
    args = ap.parse_args()

    with zipfile.ZipFile(args.src) as z:
        names = z.namelist()
        datamap = {n: z.read(n) for n in names}
    hroot = etree.fromstring(datamap["Contents/header.xml"])
    sroot = etree.fromstring(datamap["Contents/section0.xml"])
    cidx = charpr_index(hroot)
    body_style = body_mode_style(hroot, sroot, cidx)
    before_text = doc_text(sroot)
    report = ["[body mode style (height,fontRef)]: %s" % (body_style,)]

    na = nb = nc = 0
    if "A" in args.rules.upper():
        na = rule_a(hroot, sroot, report)
    if "B" in args.rules.upper():
        nb = rule_b(hroot, sroot, cidx, body_style, report)
    if args.flag_c:
        nc = rule_c_flag(sroot, cidx, body_style, report)

    datamap["Contents/header.xml"] = etree.tostring(hroot, xml_declaration=True, encoding="UTF-8", standalone=True)
    datamap["Contents/section0.xml"] = etree.tostring(sroot, xml_declaration=True, encoding="UTF-8", standalone=True)
    repack(args.src, args.dst, datamap)

    # ---- self-verify ----
    after_text = doc_text(etree.fromstring(datamap["Contents/section0.xml"]))
    text_ok = (before_text == after_text)   # R-A/R-B/R-C do not change any character text
    zip_ok = zipfile.ZipFile(args.dst).testzip() is None
    try:
        from hwpx.document import HwpxDocument
        d = HwpxDocument.open(args.dst); npar = sum(1 for _ in d.paragraphs); open_ok = npar > 0
    except Exception as e:
        open_ok = False; npar = -1; report.append("  [warn] hwpx open: %r" % e)

    print("\n".join(report))
    print("\n=== SUMMARY ===")
    print("R-A intent normalized:", na, "| R-B bracket bolded:", nb, "| R-C candidates flagged:", nc)
    print("TEXT-INVARIANT:", text_ok, "| ZIP-OK:", zip_ok, "| HWPX-OPEN:", open_ok, "(paras=%s)" % npar)
    print("VERDICT:", "PASS" if (text_ok and zip_ok and open_ok) else "FAIL")

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
