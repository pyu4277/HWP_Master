"""Reusable OWPML in-place fill helpers for the 전기기초실습 서식2~5 target.
Style-safe: only duplicates the target's OWN elements (no cross-file splice).
"""
import zipfile, copy
from lxml import etree

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
def q(tag): return "{%s}%s" % (HP, tag)
def local(e): return etree.QName(e.tag).localname
def ptext(p): return "".join(t.text or "" for t in p.iter() if local(t) == "t")

def load_section(path, name="Contents/section0.xml"):
    with zipfile.ZipFile(path) as z:
        xml = z.read(name)
    return etree.fromstring(xml)

def serialize(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

def repack(src_hwpx, dst_hwpx, new_section_bytes, section_name="Contents/section0.xml"):
    with zipfile.ZipFile(src_hwpx) as zin:
        names = zin.namelist()
        datamap = {n: zin.read(n) for n in names}
    datamap[section_name] = new_section_bytes
    with zipfile.ZipFile(dst_hwpx, "w", zipfile.ZIP_DEFLATED) as zout:
        # mimetype first, stored
        if "mimetype" in datamap:
            zi = zipfile.ZipInfo("mimetype"); zi.compress_type = zipfile.ZIP_STORED
            zout.writestr(zi, datamap["mimetype"])
        for n in names:
            if n == "mimetype":
                continue
            zout.writestr(n, datamap[n])

# ---- paragraph helpers ----
def set_para_text(p, text):
    """Set the paragraph's first run <t> to text; clear other runs' <t>; strip linesegarray."""
    runs = p.findall(q("run"))
    if not runs:
        run = etree.SubElement(p, q("run"))
        runs = [run]
    first = runs[0]
    t = first.find(q("t"))
    if t is None:
        t = etree.SubElement(first, q("t"))
    t.text = text
    for r in runs[1:]:
        for tt in r.findall(q("t")):
            r.remove(tt)
    for ls in p.findall(q("linesegarray")):
        p.remove(ls)
    return p

def set_para_pr(p, ppid):
    """Set the paragraph's paraPrIDRef (changes indent + auto-bullet style)."""
    if ppid is not None:
        p.set("paraPrIDRef", str(ppid))

def dup_para(p, text, ppid=None):
    """Deepcopy paragraph p, set text, optionally change paraPrIDRef. Return new paragraph."""
    newp = copy.deepcopy(p)
    set_para_text(newp, text)
    set_para_pr(newp, ppid)
    return newp

def replace_para_block(anchor_p, lines, ppid=None):
    """Given an anchor paragraph, set it to lines[0] and insert lines[1:] as following siblings.
    If ppid given, all paragraphs get that paraPrIDRef (e.g. an auto-bullet style). Returns list."""
    result = []
    set_para_text(anchor_p, lines[0])
    set_para_pr(anchor_p, ppid)
    result.append(anchor_p)
    prev = anchor_p
    for ln in lines[1:]:
        np = dup_para(anchor_p, ln, ppid)
        prev.addnext(np)
        prev = np
        result.append(np)
    return result

# ---- finders ----
def find_paras(root):
    return [e for e in root.iter() if local(e) == "p"]

def find_para_exact(root, text):
    for p in find_paras(root):
        if ptext(p).strip() == text.strip():
            return p
    return None

def find_next_sibling_para(p, match=None):
    parent = p.getparent()
    sibs = list(parent)
    idx = sibs.index(p)
    for e in sibs[idx+1:]:
        if local(e) == "p":
            if match is None or ptext(e).strip() == match:
                return e
    return None

def find_tables(root):
    return [e for e in root.iter() if local(e) == "tbl"]

def table_first_cell_text(tbl):
    tc = tbl.find("./{*}tr/{*}tc")
    return ptext(tc).strip() if tc is not None else ""

def find_table_by_headers(root, first_cell, second_cell=None):
    for t in find_tables(root):
        rows = t.findall("./{*}tr")
        if not rows: continue
        cells = rows[0].findall("./{*}tc")
        c0 = ptext(cells[0]).strip() if cells else ""
        if c0 == first_cell:
            if second_cell is None:
                return t
            c1 = ptext(cells[1]).strip() if len(cells) > 1 else ""
            if c1 == second_cell:
                return t
    return None

def cell_at(tbl, row, col):
    """Return the tc element at logical (row,col) using tr/tc order (col = tc index in that tr)."""
    rows = tbl.findall("./{*}tr")
    if row >= len(rows): return None
    cells = rows[row].findall("./{*}tc")
    if col >= len(cells): return None
    return cells[col]

def cell_paras(tc):
    subList = tc.find("./{*}subList")
    if subList is None:
        return []
    return subList.findall("./{*}p")

def _renumber_row_addr(tbl):
    """Set each tc's cellAddr@rowAddr to its tr index."""
    for ri, tr in enumerate(tbl.findall("./{*}tr")):
        for tc in tr.findall("./{*}tc"):
            ca = tc.find("./{*}cellAddr")
            if ca is not None:
                ca.set("rowAddr", str(ri))

def add_table_row(tbl):
    """Duplicate the last <tr>, clear its cell text, append, fix rowCnt & rowAddr. Return new tr."""
    trs = tbl.findall("./{*}tr")
    last = trs[-1]
    newtr = copy.deepcopy(last)
    # clear text in each cell
    for tc in newtr.findall("./{*}tc"):
        for p in tc.findall(".//{*}p"):
            set_para_text(p, "")
        # keep only first paragraph in each cell subList
        sl = tc.find("./{*}subList")
        if sl is not None:
            ps = sl.findall("./{*}p")
            for extra in ps[1:]:
                sl.remove(extra)
    last.addnext(newtr)
    # bump rowCnt
    rc = tbl.get("rowCnt")
    if rc is not None:
        tbl.set("rowCnt", str(int(rc) + 1))
    _renumber_row_addr(tbl)
    return newtr

def remove_last_table_row(tbl):
    trs = tbl.findall("./{*}tr")
    last = trs[-1]
    last.getparent().remove(last)
    rc = tbl.get("rowCnt")
    if rc is not None:
        tbl.set("rowCnt", str(int(rc) - 1))
    _renumber_row_addr(tbl)

def remove_pics(elem):
    """Remove all <hp:pic> (and picture/image variants) descendants; return count."""
    n = 0
    for pic in list(elem.iter()):
        if local(pic) in ("pic", "picture", "image"):
            parent = pic.getparent()
            if parent is not None:
                parent.remove(pic)
                n += 1
    return n

def remove_pics_by_ref(root, refs):
    """Remove <hp:pic> (floating image controls) whose binaryItemIDRef is in refs. Returns count."""
    refs = set(refs)
    n = 0
    for pic in list(root.iter()):
        if local(pic) != "pic":
            continue
        hit = any(c.get("binaryItemIDRef") in refs for c in pic.iter())
        if hit:
            parent = pic.getparent()
            if parent is not None:
                parent.remove(pic)
                n += 1
    return n

def replace_in_runs(root, old, new):
    """Substring replace within each run's <t> text across the document. Returns count."""
    n = 0
    for r in root.iter(q("run")):
        for t in r.findall(q("t")):
            if t.text and old in t.text:
                t.text = t.text.replace(old, new)
                n += 1
    return n

def set_cell_lines(tc, lines, ppid=None):
    """Set a cell's content to the given lines (one <p> per line). Reuses/clones the cell's first p.
    If ppid given, each paragraph gets that paraPrIDRef (auto-bullet)."""
    subList = tc.find("./{*}subList")
    ps = subList.findall("./{*}p")
    base = ps[0]
    if not lines:
        lines = [""]
    set_para_text(base, lines[0])
    set_para_pr(base, ppid)
    for extra in ps[1:]:
        subList.remove(extra)
    prev = base
    for ln in lines[1:]:
        np = dup_para(base, ln, ppid)
        prev.addnext(np)
        prev = np
    return tc
