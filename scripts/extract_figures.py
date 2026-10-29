#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inspect a paper PDF and extract figures/tables as high-resolution PNGs.

Examples:
  python extract_figures.py paper.pdf list
  python extract_figures.py paper.pdf text --pages 1-8
  python extract_figures.py paper.pdf page --page 5 --dpi 200 --outdir pages
  python extract_figures.py paper.pdf auto --caption "Fig. 2" --out fig2.png
  python extract_figures.py paper.pdf crop --page 5 --relative 0.05,0.18,0.95,0.55 --out fig2.png
"""
import argparse
import re
import sys
from pathlib import Path

try:
    import pymupdf
except ImportError:
    try:
        import fitz as pymupdf  # older PyMuPDF
    except ImportError:
        print("ERROR: PyMuPDF is required. Install with: pip install pymupdf")
        sys.exit(1)


CAPTION_RE = re.compile(
    r"^\s*(?:(Fig(?:ure)?)\s*(\d+)\s*[.:]|(TABLE|Table)\s*([0-9]+|[IVXLC]+)\b)"
)


def compact(text, limit=None):
    text = re.sub(r"\s+", " ", text or "").strip()
    return text[:limit] + ("..." if limit and len(text) > limit else "")


def norm_label(kind, num):
    kind = "Fig." if kind.lower().startswith("fig") else "Table"
    return f"{kind} {num}"


def page_text_blocks(page):
    blocks = []
    for b in page.get_text("blocks", sort=True):
        if len(b) >= 7 and b[6] != 0:
            continue
        x0, y0, x1, y1, text = b[:5]
        if not text or not text.strip():
            continue
        blocks.append({
            "rect": pymupdf.Rect(x0, y0, x1, y1),
            "text": text.strip(),
        })
    return blocks


def iter_captions(doc, pages=None):
    found = []
    for pno in pages or range(len(doc)):
        page = doc[pno]
        blocks = page_text_blocks(page)
        for b in blocks:
            first_line = b["text"].splitlines()[0] if b["text"] else ""
            m = CAPTION_RE.match(first_line)
            if not m:
                continue
            if m.group(1):
                label = norm_label(m.group(1), m.group(2))
                kind = "fig"
            else:
                label = norm_label(m.group(3), m.group(4))
                kind = "table"
            found.append({
                "page": pno + 1,
                "label": label,
                "kind": kind,
                "caption": compact(b["text"], 300),
                "rect": b["rect"],
            })
    return found


def parse_pages(spec, total):
    if not spec:
        return None
    pages = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            pages.update(range(int(a), int(b) + 1))
        else:
            pages.add(int(part))
    return sorted(p - 1 for p in pages if 1 <= p <= total)


def two_column_mid(page, blocks):
    width = page.rect.width
    mid = width / 2
    left = sum(1 for b in blocks if b["rect"].x1 < mid - 8)
    right = sum(1 for b in blocks if b["rect"].x0 > mid + 8)
    wide = sum(1 for b in blocks if b["rect"].width > width * 0.65)
    if left >= 3 and right >= 3 and wide < len(blocks) * 0.5:
        return mid
    return None


def column_bounds(page, cap_rect, blocks):
    width = page.rect.width
    mid = two_column_mid(page, blocks)
    if mid is None or cap_rect.width > width * 0.55:
        return 0.0, width
    if (cap_rect.x0 + cap_rect.x1) / 2 < mid:
        return 0.0, mid + 2
    return mid - 2, width


def area(rect):
    return max(0.0, rect.width) * max(0.0, rect.height)


def intersects(rect, col):
    return min(rect.x1, col[1]) - max(rect.x0, col[0]) > 12


def graphic_candidates(page, col, above_rect=None, below_rect=None):
    """Return (image rects, drawing rects) that sit in the requested band."""
    imgs, draws = [], []
    for info in page.get_image_info():
        r = pymupdf.Rect(info["bbox"])
        if not intersects(r, col) or area(r) < 900:
            continue
        if above_rect is not None and r.y1 <= above_rect.y0 + 3 and r.y0 >= above_rect.y0 - 0.62 * page.rect.height:
            imgs.append(r)
        if below_rect is not None and r.y0 >= below_rect.y1 - 3 and r.y1 <= below_rect.y1 + 0.62 * page.rect.height:
            imgs.append(r)
    for d in page.get_drawings():
        r = pymupdf.Rect(d["rect"])
        if not intersects(r, col) or area(r) < 16:
            continue
        if r.width > page.rect.width * 0.92 and r.height > page.rect.height * 0.92:
            continue  # page border
        if above_rect is not None and r.y1 <= above_rect.y0 + 3 and r.y0 >= above_rect.y0 - 0.62 * page.rect.height:
            draws.append(r)
        if below_rect is not None and r.y0 >= below_rect.y1 - 3 and r.y1 <= below_rect.y1 + 0.62 * page.rect.height:
            draws.append(r)
    return imgs, draws


def nearest_band(rects, anchor_y, direction, tolerance):
    """Pick rects close to the caption and union them."""
    if not rects:
        return None
    if direction == "above":
        close = [r for r in rects if anchor_y - r.y1 <= tolerance]
    else:
        close = [r for r in rects if r.y0 - anchor_y <= tolerance]
    if not close:
        close = rects
    u = close[0]
    for r in close[1:]:
        u |= r
    return u


def choose_region(page, caption, blocks):
    cap = caption["rect"]
    col = column_bounds(page, cap, blocks)
    W, H = page.rect.width, page.rect.height
    tolerance = 0.10 * H

    imgs_above, draw_above = graphic_candidates(page, col, above_rect=cap)
    imgs_below, draw_below = graphic_candidates(page, col, below_rect=cap)

    if caption["kind"] == "table":
        side = "below"
        chosen_img = nearest_band(imgs_below, cap.y1, "below", tolerance)
        chosen_draw = nearest_band(draw_below, cap.y1, "below", tolerance)
        region = None
        if chosen_draw is not None:
            region = pymupdf.Rect(col[0], cap.y1 + 1, col[1], chosen_draw.y1 + 8)
        elif chosen_img is not None:
            region = pymupdf.Rect(col[0], cap.y1 + 1, col[1], chosen_img.y1 + 8)
        else:
            region = pymupdf.Rect(col[0], cap.y1 + 1, col[1], min(H, cap.y1 + 0.30 * H))
        return region, col, side

    # Figure captions: use whichever side has the stronger graphic signal.
    above_img = nearest_band(imgs_above, cap.y0, "above", tolerance)
    below_img = nearest_band(imgs_below, cap.y1, "below", tolerance)
    above_draw = nearest_band(draw_above, cap.y0, "above", tolerance)
    below_draw = nearest_band(draw_below, cap.y1, "below", tolerance)

    def score(img, draw):
        if img is not None:
            return 3.0 * area(img)
        if draw is not None:
            return area(draw)
        return 0.0

    if score(above_img, above_draw) >= score(below_img, below_draw):
        side = "above"
        band = above_img or above_draw
        if band is not None:
            top = max(0, band.y0 - 6)
            bottom = cap.y0 - 1
        else:
            prev = [b for b in blocks if b["rect"].y1 <= cap.y0 + 2 and intersects(b["rect"], col)]
            prev_bottom = max((b["rect"].y1 for b in prev), default=cap.y0 - 0.32 * H)
            top = max(0, min(prev_bottom + 2, cap.y0 - 30))
            bottom = cap.y0 - 1
    else:
        side = "below"
        band = below_img or below_draw
        if band is not None:
            top = cap.y1 + 1
            bottom = min(H, band.y1 + 6)
        else:
            nxt = [b for b in blocks if b["rect"].y0 >= cap.y1 - 2 and intersects(b["rect"], col)]
            nxt_top = min((b["rect"].y0 for b in nxt), default=cap.y1 + 0.32 * H)
            top = cap.y1 + 1
            bottom = max(top + 30, min(H, nxt_top - 2))
    return pymupdf.Rect(col[0], top, col[1], bottom), col, side


def save_crop(page, rect, out, dpi):
    rect = pymupdf.Rect(rect) & page.rect
    if rect.is_empty or rect.width < 20 or rect.height < 20:
        raise ValueError("crop region is empty or too small")
    pix = page.get_pixmap(clip=rect, dpi=dpi, alpha=False)
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    pix.save(str(out))
    print(f"SAVED {out}  {pix.width}x{pix.height}px  clip=({rect.x0:.1f},{rect.y0:.1f},{rect.x1:.1f},{rect.y1:.1f})  dpi={dpi}")


def cmd_list(doc, args):
    pages = parse_pages(args.pages, len(doc))
    caps = iter_captions(doc, pages)
    if not caps:
        print("NO CAPTIONS FOUND.")
        print("The PDF may have no text layer (scanned/outlined). Use `page` renders and `crop` instead.")
        return
    current = None
    for c in caps:
        if c["page"] != current:
            current = c["page"]
            page = doc[current - 1]
            print(f"\nPAGE {current}  ({page.rect.width:.0f}x{page.rect.height:.0f}pt)  "
                  f"images={len(page.get_image_info())} drawings={len(page.get_drawings())}")
        r = c["rect"]
        print(f"  [{c['label']}] {c['caption']}")
        print(f"      caption_bbox=({r.x0:.0f},{r.y0:.0f},{r.x1:.0f},{r.y1:.0f}) kind={c['kind']}")
    print("\nNext: `auto --caption \"Fig. 1\" --out fig1.png`, then view the PNG and re-crop if needed.")


def cmd_text(doc, args):
    pages = parse_pages(args.pages, len(doc)) or range(len(doc))
    budget = args.max_chars
    emitted = 0
    for pno in pages:
        page = doc[pno]
        blocks = page_text_blocks(page)
        header = f"\n===== PAGE {pno + 1} ====="
        print(header)
        emitted += len(header)
        if not blocks:
            print("(no text layer)")
            continue
        for b in blocks:
            line = f"[{b['rect'].x0:.0f},{b['rect'].y0:.0f},{b['rect'].x1:.0f},{b['rect'].y1:.0f}] {compact(b['text'], 1200)}"
            if emitted + len(line) > budget:
                print(f"\n... output truncated at {budget} chars; use --pages to narrow the range.")
                return
            print(line)
            emitted += len(line)


def cmd_page(doc, args):
    pages = parse_pages(args.pages, len(doc)) or [args.page - 1]
    outdir = Path(args.outdir or ".")
    outdir.mkdir(parents=True, exist_ok=True)
    for pno in pages:
        out = outdir / (args.out if args.out and len(pages) == 1 else f"page{pno + 1}.png")
        pix = doc[pno].get_pixmap(dpi=args.dpi, alpha=False)
        pix.save(str(out))
        print(f"SAVED {out}  {pix.width}x{pix.height}px")


def cmd_auto(doc, args):
    pages = parse_pages(args.pages, len(doc)) or range(len(doc))
    target = re.sub(r"\s+", " ", args.caption.strip().lower())
    target = target.replace("figure", "fig.").replace("table ", "table ")
    target = re.sub(r"^fig\s*", "fig. ", target)
    target = re.sub(r"\s+", " ", target)
    candidates = iter_captions(doc, pages)
    match = None
    for c in candidates:
        if c["label"].lower().replace(" ", "") == target.replace(" ", ""):
            match = c
            break
    if match is None:
        labels = ", ".join(c["label"] for c in candidates) or "(none)"
        print(f"ERROR: caption {args.caption!r} not found. Available: {labels}")
        print("If this PDF has no text layer, use `page` + `crop --relative`.")
        sys.exit(2)
    page = doc[match["page"] - 1]
    blocks = page_text_blocks(page)
    region, col, side = choose_region(page, match, blocks)
    pad = args.padding
    region = pymupdf.Rect(region.x0 + pad, region.y0 + pad, region.x1 - pad, region.y1 - pad)
    print(f"MATCH {match['label']} on page {match['page']}: {match['caption']}")
    print(f"side={side} column=({col[0]:.0f},{col[1]:.0f})")
    save_crop(page, region, args.out, args.dpi)


def cmd_crop(doc, args):
    pages = parse_pages(args.pages, len(doc)) or [args.page - 1]
    page = doc[pages[0]]
    vals = [float(v) for v in args.clip.split(",")] if args.clip else [float(v) for v in args.relative.split(",")]
    if len(vals) != 4:
        raise ValueError("clip/relative needs four values: x0,y0,x1,y1")
    if args.relative:
        W, H = page.rect.width, page.rect.height
        rect = pymupdf.Rect(vals[0] * W, vals[1] * H, vals[2] * W, vals[3] * H)
    else:
        rect = pymupdf.Rect(*vals)
    save_crop(page, rect, args.out, args.dpi)


def cmd_images(doc, args):
    pages = parse_pages(args.pages, len(doc)) or range(len(doc))
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    count = 0
    for pno in pages:
        page = doc[pno]
        for info in page.get_image_info(xrefs=True):
            w, h = info.get("width", 0), info.get("height", 0)
            if w < args.min_px or h < args.min_px:
                continue
            xref = info.get("xref", 0)
            if not xref:
                continue
            try:
                img = doc.extract_image(xref)
            except Exception:
                continue
            ext = img.get("ext", "png")
            out = outdir / f"p{pno + 1}_x{xref}_{w}x{h}.{ext}"
            out.write_bytes(img["image"])
            bbox = tuple(round(v) for v in info["bbox"])
            print(f"SAVED {out}  bbox={bbox}")
            count += 1
    print(f"{count} image(s) extracted.")


def main():
    ap = argparse.ArgumentParser(description="Inspect and crop figures/tables from a paper PDF.")
    ap.add_argument("pdf")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("list", help="list figure/table captions with bounding boxes")
    p.add_argument("--pages")
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("text", help="dump text blocks with coordinates")
    p.add_argument("--pages")
    p.add_argument("--max-chars", type=int, default=60000)
    p.set_defaults(func=cmd_text)

    p = sub.add_parser("page", help="render full page(s) to PNG")
    p.add_argument("--page", type=int, default=1)
    p.add_argument("--pages")
    p.add_argument("--dpi", type=int, default=180)
    p.add_argument("--out")
    p.add_argument("--outdir", default="pages")
    p.set_defaults(func=cmd_page)

    p = sub.add_parser("auto", help="auto-crop by caption label (e.g. 'Fig. 2', 'Table I')")
    p.add_argument("--caption", required=True)
    p.add_argument("--pages")
    p.add_argument("--out", required=True)
    p.add_argument("--dpi", type=int, default=300)
    p.add_argument("--padding", type=float, default=4)
    p.set_defaults(func=cmd_auto)

    p = sub.add_parser("crop", help="crop by absolute points or relative fractions")
    p.add_argument("--page", type=int, default=1)
    p.add_argument("--pages")
    p.add_argument("--clip", help="x0,y0,x1,y1 in PDF points")
    p.add_argument("--relative", help="x0,y0,x1,y1 as fractions of page size")
    p.add_argument("--out", required=True)
    p.add_argument("--dpi", type=int, default=300)
    p.set_defaults(func=cmd_crop)

    p = sub.add_parser("images", help="extract embedded raster images")
    p.add_argument("--pages")
    p.add_argument("--outdir", default="images")
    p.add_argument("--min-px", type=int, default=160)
    p.set_defaults(func=cmd_images)

    args = ap.parse_args()
    try:
        doc = pymupdf.open(args.pdf)
        if doc.is_encrypted:
            doc.authenticate("")
        args.func(doc, args)
    except Exception as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
