# -*- coding: utf-8 -*-
"""
harness/vision.py - A4 비전 검증 어댑터 (LOCAL, COM 렌더 정본).

수정 HWPX 를 COM 으로 PDF 렌더 후 PNG 로 변환하고, 객관적 백본(단절 실측)을 함께 낸다.
비전(의미 판정: 글머리 정렬/괄호 bold 육안 확인)은 하네스의 비전 에이전트(모델)가 PNG 를 보고 수행한다.
이 모듈은 그 에이전트에 (a) 렌더 PNG, (b) 단절 실측 리포트를 공급한다.

**로컬 전용**: render 는 win32com(Hancom Office 2024) 필요. fitz(PyMuPDF)는 크로스플랫폼.
클라우드에서 import 는 되지만 render_* 호출은 COM 부재로 실패한다(의도된 경계).

CLI: python harness/vision.py <fixed.hwpx> [--dpi 150] [--outdir tests/_work/png]
"""
import sys, os, argparse

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
import hwpx_linefix as LF  # render_pdf / visual_lines / detect_splits (COM 부분은 lazy import)


def render_pdf(hwpx, pdf_path):
    """COM 렌더(로컬 전용). fresh-Dispatch 패턴은 linefix.render_pdf 재사용(E58 준수)."""
    return LF.render_pdf(os.path.abspath(hwpx), os.path.abspath(pdf_path))


def pdf_to_pngs(pdf_path, outdir, dpi=150):
    """PDF -> 페이지별 PNG. fitz 만 사용(크로스플랫폼)."""
    import fitz
    os.makedirs(outdir, exist_ok=True)
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    d = fitz.open(pdf_path)
    pngs = []
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    for pno in range(len(d)):
        pix = d[pno].get_pixmap(matrix=mat)
        out = os.path.join(outdir, "%s_p%03d.png" % (base, pno))
        pix.save(out)
        pngs.append(out)
    d.close()
    return pngs


def render_pngs(hwpx, outdir, dpi=150):
    """HWPX -> (COM PDF) -> PNG 목록. 로컬 전용."""
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(hwpx))[0] + ".pdf")
    render_pdf(hwpx, pdf)
    return pdf_to_pngs(pdf, outdir, dpi=dpi)


def split_report(hwpx, workdir=None):
    """객관적 백본: 렌더 후 단절 어절 실측(A4 비전 판정의 정본 근거).
    returns {"count": n, "splits": [{"head_line","tail_line"}...], "pages": npages}."""
    workdir = workdir or os.path.dirname(os.path.abspath(hwpx))
    pdf = os.path.join(workdir, "_vision_render.pdf")
    render_pdf(hwpx, pdf)
    vl = LF.visual_lines(pdf)
    sp = LF.detect_splits(vl)
    pages = 1 + max((v["page"] for v in vl), default=0)
    return {"count": len(sp),
            "splits": [{"head_line": s["head_line"][-24:], "tail_line": s["tail_line"][:24]} for s in sp],
            "pages": pages}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("hwpx")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--outdir", default=None)
    args = ap.parse_args()
    outdir = args.outdir or os.path.join(os.path.dirname(os.path.abspath(args.hwpx)), "png")
    rep = split_report(args.hwpx)
    print("[vision] splits=%d pages=%d" % (rep["count"], rep["pages"]))
    for s in rep["splits"]:
        print("   split: ...%s | %s..." % (s["head_line"], s["tail_line"]))
    pngs = render_pngs(args.hwpx, outdir, dpi=args.dpi)
    print("[vision] rendered %d PNG -> %s" % (len(pngs), outdir))
    for p in pngs:
        print("   ", p)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
