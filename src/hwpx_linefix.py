# -*- coding: utf-8 -*-
"""
hwpx_linefix.py - R-D 정밀 자간 단절방지 (COM 렌더-인-루프 엔진).

단절 어절(줄 끝에서 어절 중간이 다음 줄로 넘어감)을 자간(narrow/widen 최적)으로 1줄화.
정본 = COM 렌더 PDF(한/글 실제 줄바꿈). linesegarray 신뢰 금지(E57).

파이프라인: HWPX --win32com--> PDF --fitz--> 시각 줄 --> 단절 탐지 --> 문단/오프셋 매핑
          --> 방향최적 자간(줄-스팬, 글머리 제외) 부분-run 적용 --> 저장 --> 재렌더 --> 반복.

로컬 전용(Hancom Office 2024 COM). 요건: win32com, fitz(PyMuPDF), lxml.
"""
import sys, os, copy, zipfile, argparse, time
from lxml import etree

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HH = "http://www.hancom.co.kr/hwpml/2011/head"
def L(e): return etree.QName(e.tag).localname
def qhp(t): return "{%s}%s" % (HP, t)
def qhh(t): return "{%s}%s" % (HH, t)
BULLET_GLYPHS = set("○●▪·•◦□■⁃∙-")

def log(msg):
    print(msg, flush=True)

# ---------------- COM render (FRESH instance per render — confirmed-clean pattern) ----------------
# NOTE: reusing one HwpObject + accessing XHwpDocuments corrupts state ("RPC 서버를 사용할 수 없습니다"
# then AttributeError('Open')). Confirmed-working pattern = fresh Dispatch + Open + SaveAs + Quit each call.
# Do NOT access XHwpDocuments; do NOT force-kill Hwp.exe mid-run (that corrupts the COM server). See E58.
def render_pdf(hwpx_abs, pdf_abs):
    import win32com.client as wc
    import pythoncom
    pythoncom.CoInitialize()
    try:
        h = wc.Dispatch("HWPFrame.HwpObject")
        try: h.RegisterModule("FilePathCheckDLL", "AutomationModule")
        except Exception: pass
        if os.path.exists(pdf_abs):
            try: os.remove(pdf_abs)
            except Exception: pass
        h.Open(hwpx_abs, "HWPX", "")
        h.SaveAs(pdf_abs, "PDF", "")
        for _ in range(60):
            if os.path.exists(pdf_abs) and os.path.getsize(pdf_abs) > 0:
                break
            time.sleep(0.1)
        try: h.Quit()
        except Exception: pass
    finally:
        pythoncom.CoUninitialize()
    return pdf_abs
def close_hwp():
    pass  # fresh-per-render: nothing to close globally

# ---------------- fitz visual lines ----------------
def visual_lines(pdf_abs):
    import fitz
    d = fitz.open(pdf_abs)
    out = []
    for pno in range(len(d)):
        dd = d[pno].get_text("dict")
        for bi, b in enumerate(dd.get("blocks", [])):
            if b.get("type", 0) != 0: continue
            for ln in b.get("lines", []):
                t = "".join(sp["text"] for sp in ln["spans"])
                if t.strip():
                    out.append({"page": pno, "block": bi, "text": t})
    d.close()
    return out

def is_kr(c): return '가' <= c <= '힣'
def detect_splits(vlines):
    """consecutive lines in same block: line ends Korean + next starts Korean + no space at boundary = split 어절."""
    sp = []
    for i in range(len(vlines) - 1):
        a, b = vlines[i], vlines[i + 1]
        if a["page"] != b["page"] or a["block"] != b["block"]: continue
        at, bt = a["text"], b["text"]
        if not at or not bt or at[-1].isspace() or bt[0].isspace(): continue
        if is_kr(at[-1]) and is_kr(bt[0]):
            sp.append({"head_line": at, "tail_line": bt})
    return sp

# ---------------- OWPML leaf paragraph text ----------------
def own_text(p):
    out = []
    def rec(e):
        for c in e:
            if L(c) == "p": continue
            if L(c) == "t": out.append(c.text or "")
            else: rec(c)
    rec(p); return "".join(out)

def leaf_paras(sroot):
    res = []
    for p in sroot.iter():
        if L(p) != "p": continue
        # leaf = has own text and does NOT contain nested <hp:p> that hold text (i.e., not a table container)
        has_nested_p = any(L(d) == "p" for r in p for d in r.iter() if d is not r)
        t = own_text(p)
        if t.strip() and not has_nested_p:
            res.append({"p": p, "own": t})
    return res

# ---------------- non-destructive integrity gates (README/SKILL: text-invariant + testzip + hwpx-open) ----
def all_text(root):
    """Every <t> text in document order, each counted once (consistent text-invariant measure)."""
    return "".join((t.text or "") for t in root.iter() if L(t) == "t")

def section_text(hwpx):
    with zipfile.ZipFile(hwpx) as z:
        return all_text(etree.fromstring(z.read("Contents/section0.xml")))

def integrity_check(hwpx, orig_text):
    """Return (ok, problems[]). Enforces: zip testzip, text-invariant, hwpx openable with paras>0."""
    problems = []
    try:
        if zipfile.ZipFile(hwpx).testzip() is not None:
            problems.append("zip-testzip-failed")
    except Exception as e:
        problems.append("zip-open:%r" % e)
    try:
        if section_text(hwpx) != orig_text:
            problems.append("text-changed")
    except Exception as e:
        problems.append("text-read:%r" % e)
    try:
        from hwpx.document import HwpxDocument
        d = HwpxDocument.open(hwpx)
        if sum(1 for _ in d.paragraphs) <= 0:
            problems.append("hwpx-open-zero-paras")
    except Exception as e:
        problems.append("hwpx-open:%r" % e)
    return (len(problems) == 0, problems)

# ---------------- map split -> paragraph + break offset ----------------
def map_split(split, leaves):
    a = split["head_line"].rstrip()
    b = split["tail_line"]
    for K in (14, 10, 7, 5):
        ha, hb = a[-K:], b[:K]
        key = ha + hb
        hits = [(lf, lf["own"].find(key)) for lf in leaves if key and lf["own"].find(key) >= 0]
        if len(hits) == 1:
            lf, idx = hits[0]
            return lf["p"], lf["own"], idx + len(ha)
    return None, None, None

# ---------------- span geometry ----------------
def line_start_offset(own, break_off, all_break_offs):
    """start offset of the visual line that ends at break_off (= previous break offset, or 0)."""
    prev = 0
    for bo in sorted(all_break_offs):
        if bo >= break_off: break
        prev = bo
    return prev

def word_bounds(own, b):
    ls = b
    while ls > 0 and not own[ls - 1].isspace(): ls -= 1
    le = b
    while le < len(own) and not own[le].isspace(): le += 1
    return ls, le

def skip_bullet(own, s):
    while s < len(own) and (own[s] in BULLET_GLYPHS or own[s].isspace()):
        s += 1
    return s

# ---------------- charPr spacing twin ----------------
def charpr_index(hroot):
    return {cp.get("id"): cp for cp in hroot.iter() if L(cp) == "charPr"}
def charpr_container(hroot):
    for cp in hroot.iter():
        if L(cp) == "charPr": return cp.getparent()
    return None
def spacing_twin(hroot, base_id, pct, cidx, cache):
    key = (base_id, pct)
    if key in cache: return cache[key]
    base = cidx.get(base_id)
    if base is None: return base_id
    newcp = copy.deepcopy(base)
    ids = [int(c.get("id")) for c in cidx.values() if c.get("id") and c.get("id").isdigit()]
    nid = str(max(ids) + 1)
    newcp.set("id", nid)
    for sp in list(newcp):
        if L(sp) == "spacing": newcp.remove(sp)
    # insert <hh:spacing> after relSz (schema: fontRef,ratio,spacing,relSz,offset,...) -> actually spacing before relSz
    order = ["fontRef", "ratio"]
    kids = list(newcp); idx = 0
    for i, c in enumerate(kids):
        if L(c) in order: idx = i + 1
    spel = etree.Element(qhh("spacing"))
    for a in ("hangul", "latin", "hanja", "japanese", "other", "symbol", "user"):
        spel.set(a, str(pct))
    newcp.insert(idx, spel)
    cont = charpr_container(hroot); cont.append(newcp)
    c = cont.get("itemCnt")
    if c is not None: cont.set("itemCnt", str(int(c) + 1))
    cidx[nid] = newcp; cache[key] = nid
    return nid

# ---------------- apply spacing to a char span within a paragraph ----------------
def apply_span_spacing(para, hroot, s, e, pct, cidx, cache):
    """Retag runs covering own-text [s,e) to a spacing twin (pct). Splits partial runs. Non-destructive."""
    if e <= s: return 0
    changed = 0
    off = 0
    # iterate direct run children in order (own runs)
    for run in list(para.findall(qhp("run"))):
        ts = run.findall(qhp("t"))
        if not ts:
            # run may carry inline objects; skip but advance offset by its text len (0 for no t)
            continue
        # assume single <t> per run for these docs; handle first t
        t = ts[0]
        txt = t.text or ""
        rstart, rend = off, off + len(txt)
        off = rend
        if rend <= s or rstart >= e or not txt.strip():
            continue
        # overlap [max(s,rstart), min(e,rend))
        a = max(s, rstart) - rstart
        b = min(e, rend) - rstart
        base = run.get("charPrIDRef")
        tw = spacing_twin(hroot, base, pct, cidx, cache)
        before, mid, after = txt[:a], txt[a:b], txt[b:]
        # rebuild: before(base) | mid(twin) | after(base)
        t.text = before
        anchor = run
        if before == "":
            # reuse run as mid
            run.set("charPrIDRef", str(tw)); t.text = mid
            anchor = run
            tail_txt = after
        else:
            mid_run = copy.deepcopy(run);
            for x in mid_run.findall(qhp("t"))[1:]: mid_run.remove(x)
            mid_run.set("charPrIDRef", str(tw)); mid_run.findall(qhp("t"))[0].text = mid
            run.addnext(mid_run); anchor = mid_run
            tail_txt = after
        if tail_txt:
            aft_run = copy.deepcopy(run)
            for x in aft_run.findall(qhp("t"))[1:]: aft_run.remove(x)
            aft_run.set("charPrIDRef", str(base)); aft_run.findall(qhp("t"))[0].text = tail_txt
            anchor.addnext(aft_run)
        changed += 1
    return changed

# ---------------- paragraph visual-line break offsets (align fitz lines to own_text) ----------------
def para_break_offsets(own, vlines_texts):
    """Given a paragraph's own_text and its ordered visual line texts, return break offsets (start of each line>0)."""
    offs = []
    cur = 0
    joined = own
    for k, lt in enumerate(vlines_texts[1:], start=1):
        # find start of this visual line in own_text after cur
        probe = lt.strip()[:12]
        if not probe: continue
        idx = joined.find(probe, cur)
        if idx < 0:
            idx = joined.find(probe[:6], cur)
        if idx < 0: continue
        offs.append(idx); cur = idx
    return offs

def para_visual_lines(para_own, all_vlines):
    """find contiguous run of visual lines whose concatenation matches para_own; return their texts."""
    probe = para_own.strip()[:12]
    if not probe: return []
    start = None
    for i, vl in enumerate(all_vlines):
        if vl["text"].strip().startswith(probe[:8]) or probe[:8] in vl["text"]:
            start = i; break
    if start is None: return []
    # gather following lines in same block until concatenated length ~ len(para_own)
    texts = []; blk = all_vlines[start]["block"]; pg = all_vlines[start]["page"]; acc = 0
    for vl in all_vlines[start:]:
        if vl["page"] != pg or vl["block"] != blk: break
        texts.append(vl["text"]); acc += len(vl["text"].strip())
        if acc >= len(para_own.strip()): break
    return texts

# ---------------- main fix loop ----------------
def fix_document(work_hwpx, workdir, max_global=40, verbose=True):
    pdf = os.path.join(workdir, "_lf_render.pdf")
    fixed_words = 0; failed = []
    # capture the document's text ONCE before any edit — the non-destructive invariant baseline.
    orig_text = section_text(work_hwpx)
    for git in range(max_global):
        render_pdf(os.path.abspath(work_hwpx), os.path.abspath(pdf))
        vl = visual_lines(pdf)
        splits = detect_splits(vl)
        log("[iter %d] visual_lines=%d splits=%d" % (git, len(vl), len(splits)))
        if not splits:
            ok, probs = integrity_check(work_hwpx, orig_text)
            log("[done] no split 어절 remain after %d iters, %d words fixed | integrity=%s %s"
                % (git, fixed_words, "PASS" if ok else "FAIL", "" if ok else probs))
            return {"clean": True, "iters": git, "fixed": fixed_words, "failed": failed,
                    "integrity_ok": ok, "integrity": probs}
        # load current OWPML
        with zipfile.ZipFile(work_hwpx) as z:
            names = z.namelist(); dm = {n: z.read(n) for n in names}
        sroot = etree.fromstring(dm["Contents/section0.xml"])
        hroot = etree.fromstring(dm["Contents/header.xml"])
        cidx = charpr_index(hroot); cache = {}
        leaves = leaf_paras(sroot)
        # take the TOP split, map, fix
        top = splits[0]
        para, own, boff = map_split(top, leaves)
        if para is None:
            log("  [skip] cannot map split: %r|%r" % (top["head_line"][-10:], top["tail_line"][:10]))
            failed.append(top);
            # try next splits by removing this one visually is hard; abort to avoid loop
            if len(splits) == 1:
                return {"clean": False, "iters": git, "fixed": fixed_words, "failed": failed, "reason": "unmappable"}
            # crude: try second split
            top = splits[1]; para, own, boff = map_split(top, leaves)
            if para is None:
                return {"clean": False, "iters": git, "fixed": fixed_words, "failed": failed, "reason": "unmappable2"}
        # geometry
        pv = para_visual_lines(own, vl)
        breaks = para_break_offsets(own, pv) if pv else [boff]
        if boff not in breaks: breaks = sorted(set(breaks + [boff]))
        lstart = line_start_offset(own, boff, breaks)
        lstart = skip_bullet(own, lstart)
        ls, le = word_bounds(own, boff)
        head_len, tail_len = boff - ls, le - boff
        if head_len < tail_len:
            direction = "widen"; span_s, span_e = lstart, ls; sign = +1
        else:
            direction = "narrow"; span_s, span_e = lstart, le; sign = -1
        if span_e <= span_s:
            span_s, span_e = lstart, le; sign = -1; direction = "narrow(fallback)"
        # magnitude search: apply increasing |pct| up to 20, re-render, re-check this word
        word = own[ls:le]
        solved = False
        for mag in (4, 8, 12, 16, 20):
            pct = sign * mag
            # snapshot pre-edit bytes so a corrupt/text-mutating edit can be rolled back non-destructively.
            with open(work_hwpx, "rb") as f: prev_bytes = f.read()
            # fresh load to apply cleanly each try
            with zipfile.ZipFile(work_hwpx) as z:
                names = z.namelist(); dm2 = {n: z.read(n) for n in names}
            sr = etree.fromstring(dm2["Contents/section0.xml"])
            hr = etree.fromstring(dm2["Contents/header.xml"])
            ci = charpr_index(hr); ca = {}
            lv = leaf_paras(sr)
            pp2, own2, boff2 = map_split(top, lv)
            if pp2 is None: break
            apply_span_spacing(pp2, hr, span_s, span_e, pct, ci, ca)
            dm2["Contents/section0.xml"] = etree.tostring(sr, xml_declaration=True, encoding="UTF-8", standalone=True)
            dm2["Contents/header.xml"] = etree.tostring(hr, xml_declaration=True, encoding="UTF-8", standalone=True)
            with zipfile.ZipFile(work_hwpx, "w", zipfile.ZIP_DEFLATED) as zo:
                if "mimetype" in dm2:
                    zi = zipfile.ZipInfo("mimetype"); zi.compress_type = zipfile.ZIP_STORED
                    zo.writestr(zi, dm2["mimetype"])
                for n in names:
                    if n == "mimetype": continue
                    zo.writestr(n, dm2[n])
            # non-destructive gate BEFORE spending a render: reject any edit that breaks zip/text/openability.
            ok, probs = integrity_check(work_hwpx, orig_text)
            if not ok:
                log("  [integrity FAIL pct=%d] %s -> rollback, skip word '%s'" % (pct, probs, word))
                with open(work_hwpx, "wb") as f: f.write(prev_bytes)
                failed.append({"word": word, "integrity": probs, "head": top["head_line"][-10:]})
                solved = False
                break
            render_pdf(os.path.abspath(work_hwpx), os.path.abspath(pdf))
            vl2 = visual_lines(pdf); sp2 = detect_splits(vl2)
            still = any((s["head_line"].rstrip()[-8:] == top["head_line"].rstrip()[-8:]) for s in sp2)
            log("  [%s pct=%d] word='%s' still_split=%s (total splits %d)" % (direction, pct, word, still, len(sp2)))
            if not still:
                solved = True; fixed_words += 1; break
        if not solved:
            log("  [fail] could not resolve '%s' within +-20; leaving and moving on" % word)
            failed.append({"word": word, "head": top["head_line"][-10:]})
            # to avoid infinite loop on same word, we accept current state and continue;
            # but since it re-detects top-first, mark by nudging: leave applied 20 and continue to next global iter
    ok, probs = integrity_check(work_hwpx, orig_text)
    return {"clean": False, "iters": max_global, "fixed": fixed_words, "failed": failed,
            "reason": "max_iter", "integrity_ok": ok, "integrity": probs}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("hwpx")
    ap.add_argument("--max", type=int, default=40)
    args = ap.parse_args()
    workdir = os.path.dirname(os.path.abspath(args.hwpx))
    t0 = time.time()
    try:
        res = fix_document(args.hwpx, workdir, max_global=args.max)
    finally:
        close_hwp()
    log("=== RESULT %s (%.1fs) ===" % (res, time.time() - t0))

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
