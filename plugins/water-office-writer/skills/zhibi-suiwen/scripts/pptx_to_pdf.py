#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pptx_to_pdf.py - PowerPoint COM 转 PDF

用法:
    python pptx_to_pdf.py --input <pptx> [--output <pdf>]

依赖:
    Windows + PowerPoint (pywin32)
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path


PP_FORMAT_PDF = 32  # PowerPoint 的 PDF 格式常量


def pptx_to_pdf(src: Path | str, dst: Path | str | None = None, quiet: bool = False) -> Path:
    """pptx -> PDF via PowerPoint COM"""
    import pythoncom
    import win32com.client

    src = Path(src)
    if not src.is_file():
        raise RuntimeError(f"找不到源 pptx：{src}")
    if dst is None:
        dst = src.with_suffix(".pdf")
    dst = Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    pp = None
    pres = None
    try:
        pp = win32com.client.DispatchEx("PowerPoint.Application")
        # PPT 不允许 DispatchEx 隐藏窗口 — 必须 Visible=True（PowerPoint 会话才能转 PDF）
        pp.Visible = True
        pp.DisplayAlerts = 2  # ppAlertsAll (避免确认弹窗)
        # 后台打开：WithWindow=False 不显示文档窗口，但 Visible=True 让应用消息循环可用
        pres = pp.Presentations.Open(
            str(src.resolve()),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )
        pres.SaveAs(str(dst.resolve()), PP_FORMAT_PDF)
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if pp is not None:
                pp.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    if not dst.is_file() or dst.stat().st_size == 0:
        raise RuntimeError(f"未生成 PDF：{dst}")

    if not quiet:
        print(f"[OK] PPTX -> PDF -> {dst} （{dst.stat().st_size // 1024} KB）")
    return dst


def main() -> int:
    ap = argparse.ArgumentParser(description="pptx -> PDF (PowerPoint COM)")
    ap.add_argument("--input", required=True, help="源 .pptx")
    ap.add_argument("--output", default=None, help="目标 .pdf")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()
    try:
        pptx_to_pdf(args.input, args.output, quiet=args.quiet)
    except Exception as e:
        print(f"[error] {type(e).__name__}: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())