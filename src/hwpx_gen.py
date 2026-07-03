# -*- coding: utf-8 -*-
"""
hwpx_gen.py - 합성 난잡 HWPX 생성기 (자기개선 하네스 A1 생성 에이전트의 도구).

표·캡션·주석(footnote)을 포함한 복잡·난잡 문서를 python-hwpx 로 합성하고, 각 문서의 4규칙
기대치(manifest)를 동반 출력한다. 민감데이터 0 — 전부 합성 문자열.

심는 것(planted):
  - R-A 대상 : 자동 글머리(heading BULLET) + 음수 첫줄 내어쓰기(intent<0), >=2 멤버 그룹.
  - R-B 대상 : 본문 선두 괄호 주제어 (...) + 뒤 본문 >=5자, 본문 최빈 스타일.
  - R-D 후보 : 긴 무공백 한글 어절(줄 끝 단절 유도). ** 단절 확정은 렌더에서만 ** (manifest 는 '후보').
함정(traps, 규칙이 건드리면 안 되는 것):
  - R-A: 양수 첫줄 들여쓰기 글머리 그룹 / 수동 글머리(리터럴 '-' 텍스트).
  - R-B: 제목(본문과 다른 큰 스타일)의 선두 괄호 / 괄호 뒤 잔여 <5자.

CLI:
  python hwpx_gen.py --n 20 --seed 42 --profile messy --out tests/_work/
  python hwpx_gen.py --n 1  --seed 1  --profile minimal-split --out tests/_work/  # 엔진 검증용 testdoc

출력물 *.hwpx 는 gitignore 대상(생성기 코드만 커밋, 문서는 로컬 재생성).
COM 불필요(순수 합성) — 클라우드/리눅스에서도 실행 가능.
"""
import sys, os, json, argparse, random, logging, warnings, copy
from lxml import etree as ET  # python-hwpx header.element 는 lxml (type hint 는 ET 지만 실체 lxml — 실측)

# python-hwpx 의 파트-탐색 정보 로그를 조용히(생성엔 무해).
logging.getLogger("hwpx").setLevel(logging.ERROR)
for _n in list(logging.root.manager.loggerDict):
    if _n.startswith("hwpx"):
        logging.getLogger(_n).setLevel(logging.ERROR)
warnings.filterwarnings("ignore")

from hwpx.document import HwpxDocument

# ---------------- 합성 콘텐츠 풀 (전부 무해한 합성 문자열) ----------------
TOPICS = ["표준화", "품질관리", "정보통신", "실습절차", "안전수칙", "측정오차", "회로해석",
          "데이터분석", "공정개선", "환경영향", "운영체계", "성능평가", "재료특성", "제어이론"]
SUBJECTS = ["개요", "목적", "범위", "정의", "절차", "결과", "고찰", "결론", "비고", "요약"]
WORDS = ["시스템은", "구성요소를", "포함하며", "사용자는", "결과값을", "확인한다", "본", "과정에서",
         "발생하는", "다양한", "변수들을", "종합적으로", "고려하여", "최종", "산출물을", "도출한다",
         "이때", "각", "단계별", "검증을", "수행하고", "필요시", "재현", "실험을", "통해", "신뢰성을",
         "확보한다", "관련", "기준에", "따라", "적절히", "조정되어야", "한다"]
# 긴 무공백 한글 어절(줄 끝 단절 유도용). 실제 단절 여부는 렌더 정본.
LONG_TOKENS = ["정보통신기술기반국제표준화작업반회의결과보고서에따르면",
               "품질관리시스템운영세부지침개정안검토위원회심의의결사항",
               "측정불확도평가및교정성적서발급절차운영관리규정세부기준",
               "공정능력지수산출방법론과통계적공정관리도구활용방안연구"]


def sentence(rng, nwords=None):
    n = nwords or rng.randint(8, 16)
    return " ".join(rng.choice(WORDS) for _ in range(n))


def rd_candidate_text(rng):
    """앞 어절 몇 개 + 긴 무공백 어절을 줄 끝 근처에 배치(단절 후보)."""
    lead = " ".join(rng.choice(WORDS) for _ in range(rng.randint(4, 7)))
    tok = rng.choice(LONG_TOKENS)
    tail = " ".join(rng.choice(WORDS) for _ in range(rng.randint(3, 6)))
    return "%s %s %s" % (lead, tok, tail)


def prefix(text, n=24):
    return text[:n]


# ---------------- 문서 빌더 ----------------
def _add_body(doc, text, counter, char_pr_id_ref=None, para_pr_id_ref=None):
    """본문 문단 추가 후 (문단객체, 전역 index) 반환.
    char_pr_id_ref 를 명시(inherit_style=False)해 앞 문단(예: 제목) 스타일 상속 누수를 차단한다."""
    p = doc.add_paragraph(text, char_pr_id_ref=char_pr_id_ref, para_pr_id_ref=para_pr_id_ref,
                          inherit_style=False)
    idx = counter[0]
    counter[0] += 1
    return p, idx


HH_NS = "http://www.hancom.co.kr/hwpml/2011/head"
HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HC_NS = "http://www.hancom.co.kr/hwpml/2011/core"


def _ln(e):
    return e.tag.rsplit("}", 1)[-1]


def _mm_to_hwpunit(mm):
    return int(round(mm * 7200 / 25.4))


def _ensure_bullet_parapr(doc, first_line_mm):
    """자동 글머리 paraPr 를 실물(한/글 제작 문서) 구조 그대로 주입해 그 id 를 반환한다.
    hwpx 2.9 에는 set_list_format/set_paragraph_format 이 없어(로컬 실측) header XML 을 직접 구성:
      heading type=BULLET idRef=<bullet> + hp:switch(case/default) 안 hh:margin > hc:intent.
    rule_a 탐지 요건(case/default 내 margin>intent)과 한/글 렌더 호환을 동시에 충족.
    first_line_mm<0 => 음수 내어쓰기(R-A 대상), >0 => 양수 들여쓰기(R-A 함정)."""
    hdr = doc.headers[0]
    root = hdr.element
    reflist = next(e for e in root.iter() if _ln(e) == "refList")
    parapros = next(e for e in root.iter() if _ln(e) == "paraProperties")
    # 1) bullet 정의(bullets 컨테이너가 없으면 paraProperties 앞에 생성 — refList 스키마 순서)
    bullets = next((e for e in root.iter() if _ln(e) == "bullets"), None)
    if bullets is None:
        bullets = ET.Element("{%s}bullets" % HH_NS, {"itemCnt": "0"})
        reflist.insert(list(reflist).index(parapros), bullets)
    bid = str(max([int(b.get("id", "0")) for b in bullets] + [0]) + 1)
    bu = ET.SubElement(bullets, "{%s}bullet" % HH_NS,
                       {"id": bid, "char": "●" if first_line_mm < 0 else "○", "useImage": "0"})
    ET.SubElement(bu, "{%s}paraHead" % HH_NS,
                  {"level": "0", "align": "LEFT", "useInstWidth": "0", "autoIndent": "1",
                   "widthAdjust": "0", "textOffsetType": "PERCENT", "textOffset": "50",
                   "numFormat": "DIGIT", "charPrIDRef": "4294967295", "checkable": "0"})
    bullets.set("itemCnt", str(len(list(bullets))))
    # 2) 기존 paraPr 를 clone 해 heading + switch(margin intent) 주입
    base = next(e for e in parapros if _ln(e) == "paraPr")
    pp = copy.deepcopy(base)
    pid = str(max(int(e.get("id", "0")) for e in parapros if _ln(e) == "paraPr") + 1)
    pp.set("id", pid)
    for ch in list(pp):  # 중복 방지: 기존 heading/switch/직속 margin 제거
        if _ln(ch) in ("heading", "switch", "margin"):
            pp.remove(ch)
    hd = ET.Element("{%s}heading" % HH_NS, {"type": "BULLET", "idRef": bid, "level": "0"})
    kids = list(pp)
    ai = next((i for i, c in enumerate(kids) if _ln(c) == "align"), -1)
    pp.insert(ai + 1, hd)  # 실물 순서: align, heading, breakSetting, ...
    iv = str(_mm_to_hwpunit(first_line_mm))
    lv = str(_mm_to_hwpunit(abs(first_line_mm)))
    sw = ET.Element("{%s}switch" % HP_NS)
    for tag, attrs in (("case", {"{%s}required-namespace" % HP_NS:
                                 "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"}),
                       ("default", {})):
        blk = ET.SubElement(sw, "{%s}%s" % (HP_NS, tag), attrs)
        mg = ET.SubElement(blk, "{%s}margin" % HH_NS)
        for nm, val in (("intent", iv), ("left", lv), ("right", "0"), ("prev", "0"), ("next", "0")):
            ET.SubElement(mg, "{%s}%s" % (HC_NS, nm), {"value": val, "unit": "HWPUNIT"})
        ET.SubElement(blk, "{%s}lineSpacing" % HH_NS,
                      {"type": "PERCENT", "value": "130", "unit": "HWPUNIT"})
    kids = list(pp)
    bi = next((i for i, c in enumerate(kids) if _ln(c) == "border"), len(kids))
    pp.insert(bi, sw)  # 실물 순서: ... autoSpacing, switch, border
    parapros.append(pp)
    c = parapros.get("itemCnt")
    if c is not None:
        parapros.set("itemCnt", str(int(c) + 1))
    hdr.mark_dirty()
    return pid


def _add_bullet_group(doc, texts, counter, first_line_mm, body_style):
    """자동 글머리 그룹을 '하나의 공유 paraPr'로 생성(rule_a 의 >=2 멤버 그룹 요건 충족).
    first_line_mm<0 => 음수 내어쓰기(R-A 대상), >0 => 양수 들여쓰기(R-A 함정)."""
    ppid = _ensure_bullet_parapr(doc, first_line_mm)
    idxs = []
    for t in texts:
        _, ix = _add_body(doc, t, counter, char_pr_id_ref=body_style, para_pr_id_ref=ppid)
        idxs.append(ix)
    return idxs


def _style_with_height(doc, height):
    """height(1/100pt) 지정 charPr id 확보. hwpx 2.9 ensure_run_style 은 size 미지원(로컬 실측)이라
    header.ensure_char_property(predicate/modifier)로 기존 charPr 를 클론해 height 만 바꾼 쌍둥이를 만든다."""
    hdr = doc.headers[0]
    el = hdr.ensure_char_property(
        predicate=lambda e: e.get("height") == str(height),
        modifier=lambda e: e.set("height", str(height)),
    )
    return el.get("id")


def build_doc(rng, profile):
    """(HwpxDocument, expect_dict) 생성."""
    doc = HwpxDocument.new()
    counter = [len(list(doc.paragraphs))]  # 기본 문단 수(보통 1: 빈 선두 문단)
    expect = {"R-A": [], "R-B": [], "R-D_candidates": [], "traps": []}

    # 본문 최빈 스타일(11pt=height1100)과 제목 스타일(24pt=height2400)을 명시적으로 분리한다.
    # (add_paragraph 는 기본적으로 앞 문단 스타일을 상속하므로, 본문엔 body_style 을 강제 지정.)
    body_style = _style_with_height(doc, 1100)   # 11pt
    title_style = _style_with_height(doc, 2400)  # 24pt

    # 제목: 본문과 다른 큰 스타일 + 선두 괄호 => R-B 함정(스타일 불일치로 배제되어야).
    ttl = "(%s) %s 실습 결과 보고서" % (rng.choice(SUBJECTS), rng.choice(TOPICS))
    _add_body(doc, ttl, counter, char_pr_id_ref=title_style)
    expect["traps"].append({"kind": "title-bracket-nonbody-style", "rule": "R-B",
                            "text_prefix": prefix(ttl), "must_not_change": True})

    if profile == "minimal-split":
        # 엔진 검증용 최소 문서: 단절 후보 1~2개만.
        for _ in range(rng.randint(1, 2)):
            t = rd_candidate_text(rng)
            _add_body(doc, t, counter, char_pr_id_ref=body_style)
            expect["R-D_candidates"].append({"text_prefix": prefix(t), "note": "confirm at render"})
        return doc, expect

    # --- 본문 선두 괄호 주제어 (R-B 대상) ---
    subj = rng.choice(SUBJECTS)
    body = "(%s) %s" % (subj, sentence(rng, rng.randint(10, 14)))
    _add_body(doc, body, counter, char_pr_id_ref=body_style)
    expect["R-B"].append({"token": "(%s)" % subj, "text_prefix": prefix(body)})

    # R-B 함정: 괄호 뒤 잔여 <5자.
    short = "(%s) 짧음" % rng.choice(SUBJECTS)
    _add_body(doc, short, counter, char_pr_id_ref=body_style)
    expect["traps"].append({"kind": "bracket-remainder-under-5", "rule": "R-B",
                            "text_prefix": prefix(short), "must_not_change": True})

    # --- 자동 글머리 그룹 (R-A 대상): 공유 paraPr(음수 내어쓰기, >=2 멤버), 긴 본문 ---
    a_texts = ["%s %s" % (sentence(rng, 6), rd_candidate_text(rng)) for _ in range(rng.randint(2, 3))]
    _add_bullet_group(doc, a_texts, counter, first_line_mm=-8.0, body_style=body_style)
    for t in a_texts:
        expect["R-A"].append({"text_prefix": prefix(t), "reason": "auto-bullet negative intent (>=2 members)"})
        expect["R-D_candidates"].append({"text_prefix": prefix(t), "note": "confirm at render (in bullet)"})

    # --- 자동 글머리 함정 (R-A): 양수 첫줄 들여쓰기 공유 그룹(불가침) ---
    p_texts = [sentence(rng, rng.randint(10, 14)) for _ in range(2)]
    _add_bullet_group(doc, p_texts, counter, first_line_mm=8.0, body_style=body_style)
    for t in p_texts:
        expect["traps"].append({"kind": "positive-indent-bullet", "rule": "R-A",
                                "text_prefix": prefix(t), "must_not_change": True})

    # --- 수동 글머리 함정 (R-A): 리터럴 '-' 텍스트, 자동 글머리 아님 ---
    man = "- %s" % sentence(rng, 8)
    _add_body(doc, man, counter, char_pr_id_ref=body_style)
    expect["traps"].append({"kind": "manual-bullet", "rule": "R-A",
                            "text_prefix": prefix(man), "must_not_change": True})

    # --- R-D 후보 본문 몇 개 더 ---
    for _ in range(rng.randint(2, 3)):
        t = rd_candidate_text(rng)
        _add_body(doc, t, counter, char_pr_id_ref=body_style)
        expect["R-D_candidates"].append({"text_prefix": prefix(t), "note": "confirm at render"})

    # --- 표 + 캡션 + 주석(난잡성). 표 헤더 = 규칙 배제 영역 ---
    cap = "표 1. %s 측정값" % rng.choice(TOPICS)
    _add_body(doc, cap, counter, char_pr_id_ref=body_style)
    expect["traps"].append({"kind": "caption", "rule": "R-A/R-B/R-D",
                            "text_prefix": prefix(cap), "must_not_change": True})
    try:
        doc.add_table(rows=3, cols=3)
    except Exception as e:
        logging.getLogger("hwpx_gen").warning("add_table skipped: %r", e)
    # 주석: 마지막 본문 문단에 각주.
    try:
        doc.add_footnote("각주: %s 관련 세부 사항 참조." % rng.choice(TOPICS))
    except Exception as e:
        logging.getLogger("hwpx_gen").warning("add_footnote skipped: %r", e)

    return doc, expect


def _finalize_traps_none(doc):
    return doc


# ---------------- 생성 드라이버 ----------------
def generate(n, seed, profile, out):
    os.makedirs(out, exist_ok=True)
    manifest = {"seed": seed, "profile": profile, "count": n, "docs": []}
    for i in range(1, n + 1):
        rng = random.Random("%s:%d:%s" % (seed, i, profile))  # 문서별 결정적 시드
        doc, expect = build_doc(rng, profile)
        fname = "doc%02d.hwpx" % i if profile != "minimal-split" or n > 1 else "testdoc.hwpx"
        fpath = os.path.join(out, fname)
        doc.save_to_path(fpath)
        # 생성물 자체 무결성(합성 결함 아님) 확인.
        rep = doc.validate()
        ok = getattr(rep, "ok", None)
        npar = sum(1 for _ in HwpxDocument.open(fpath).paragraphs)
        manifest["docs"].append({"file": fname, "paras": npar,
                                 "validate_ok": bool(ok) if ok is not None else "n/a",
                                 "expect": expect})
        print("[gen] %s paras=%d validate_ok=%s R-A=%d R-B=%d R-D_cand=%d traps=%d"
              % (fname, npar, ok, len(expect["R-A"]), len(expect["R-B"]),
                 len(expect["R-D_candidates"]), len(expect["traps"])), flush=True)
    mpath = os.path.join(out, "manifest.json")
    with open(mpath, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print("[gen] wrote %s (%d docs)" % (mpath, n), flush=True)
    return manifest


def main():
    ap = argparse.ArgumentParser(description="합성 난잡 HWPX 생성기(+manifest)")
    ap.add_argument("--n", type=int, default=20)
    ap.add_argument("--seed", default="42")
    ap.add_argument("--profile", choices=["messy", "minimal-split"], default="messy")
    ap.add_argument("--out", default="tests/_work/")
    args = ap.parse_args()
    generate(args.n, args.seed, args.profile, args.out)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
