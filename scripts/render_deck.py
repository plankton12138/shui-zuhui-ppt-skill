#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Render a .pptx to per-slide PNGs so the layout can be eyeballed.

Windows + PowerPoint only (uses COM). Falls back to LibreOffice if available.

  python render_deck.py out.pptx --outdir preview
"""
import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def render_powerpoint(pptx_path, outdir):
    import win32com.client  # noqa: F401

    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = None
    try:
        pres = app.Presentations.Open(str(Path(pptx_path).resolve()), WithWindow=False)
        # ppSaveAsPNG = 18 exports every slide as a numbered PNG into a folder
        pres.SaveAs(str(outdir.resolve()), 18)
    finally:
        if pres is not None:
            pres.Close()
        app.Quit()
    return sorted({p.resolve(): p for p in outdir.iterdir() if p.suffix.lower() == ".png"}.values(),
                  key=lambda p: p.name)


def render_libreoffice(pptx_path, outdir):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("Neither PowerPoint COM nor LibreOffice is available.")
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    subprocess.run([soffice, "--headless", "--convert-to", "png",
                    "--outdir", str(outdir), str(pptx_path)], check=True)
    return sorted(outdir.glob("*.png"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--outdir", default="preview")
    args = ap.parse_args()

    try:
        files = render_powerpoint(args.pptx, args.outdir)
    except Exception as exc:  # pragma: no cover - environment dependent
        print(f"PowerPoint COM failed ({exc}); trying LibreOffice...")
        files = render_libreoffice(args.pptx, args.outdir)
    for f in files:
        print(f"  {f}")
    print(f"rendered {len(files)} slide image(s) -> {args.outdir}")


if __name__ == "__main__":
    main()
