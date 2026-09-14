#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf_export.py — docx -> PDF 导出（补齐基线缺失的 PDF 输出能力）

背景：
    基线测试发现本技能只有 docx 落版，没有 PDF 输出脚本，而公文报送
    场景常需要 PDF。本模块提供两条转换通道，按可靠性降级：

        通道 1（首选）Word COM —— win32com + Microsoft Word
                      版式保真度最高（页脚 PAGE 域、三线表框线均正确渲染）
        通道 2（兜底）docx2pdf —— 内部同样调用 Word；若 COM 直连失败则回退

    两条通道都依赖 Windows + 已安装 Microsoft Word。若两者都不可用，
    本模块抛出明确错误，并提示改用 LibreOffice 无头模式替代。

用法（供 md_to_gongwen_docx.py --pdf / generate_report.py --pdf 调用）:
    from pdf_export import docx_to_pdf
    docx_to_pdf(Path("out.docx"))              # 同名 .pdf
    docx_to_pdf(Path("out.docx"), Path("x.pdf"))

命令行:
    python pdf_export.py --input out.docx [--output out.pdf] [--keep-docx]

依赖:
    pip install pywin32 docx2pdf    # 仅 Windows
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
from pathlib import Path

WD_FORMAT_PDF = 17          # Word 的 PDF 固定格式常量
WD_EXPORT_ALL = 0
WD_EXPORT_DOCUMENT_CONTENT = 0


class PdfExportError(RuntimeError):
    """PDF 导出失败（Word 不可用 / COM 异常 / 输出缺失）"""


def _abs(p: Path) -> str:
    """Word COM 需要绝对路径，且必须是 Windows 反斜杠形式"""
    return str(Path(p).resolve())


def _try_word_com(src: Path, dst: Path) -> str:
    """通道 1：Word COM 直连"""
    import pythoncom  # pywin32
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        # DispatchEx 起独立进程，避免干扰用户正在编辑的 Word 实例
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(_abs(src), ReadOnly=True, AddToRecentFiles=False)
        try:
            # ExportAsFixedFormat 比 SaveAs 更可控（不走"另存为"对话框逻辑）
            doc.ExportAsFixedFormat(
                OutputFileName=_abs(dst),
                ExportFormat=WD_FORMAT_PDF,
                OpenAfterExport=False,
                OptimizeFor=0,                     # wdExportOptimizeForPrint
                Range=WD_EXPORT_ALL,
                Item=WD_EXPORT_DOCUMENT_CONTENT,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=1,                 # 用标题生成书签
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )
        except Exception:
            # 老版本 Word 没有 ExportAsFixedFormat 的某些参数，退回 SaveAs
            doc.SaveAs(_abs(dst), FileFormat=WD_FORMAT_PDF)
        return "word-com"
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _try_docx2pdf(src: Path, dst: Path) -> str:
    """通道 2：docx2pdf 兜底"""
    from docx2pdf import convert

    convert(_abs(src), _abs(dst))
    return "docx2pdf"


def ensure_word_available() -> bool:
    """探测本机是否有 Microsoft Word（COM 可注册）"""
    try:
        import win32com.client  # noqa: F401
    except Exception:
        return False
    for exe in (
        r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
    ):
        if os.path.exists(exe):
            return True
    found = shutil.which("WINWORD.EXE")
    return bool(found)


def docx_to_pdf(src: Path | str, dst: Path | str | None = None,
                quiet: bool = False) -> Path:
    """
    docx -> PDF。返回生成的 PDF 路径。

    src : 源 .docx
    dst : 目标 .pdf（省略则与 src 同名同目录）
    """
    src = Path(src)
    if not src.is_file():
        raise PdfExportError(f"找不到源 docx：{src}")
    if dst is None:
        dst = src.with_suffix(".pdf")
    dst = Path(dst)
    dst.parent.mkdir(parents=True, exist_ok=True)

    errors: list[str] = []
    for name, fn in (("word-com", _try_word_com), ("docx2pdf", _try_docx2pdf)):
        try:
            used = fn(src, dst)
        except Exception as e:                      # noqa: BLE001
            errors.append(f"{name}: {type(e).__name__}: {e}")
            continue
        if dst.is_file() and dst.stat().st_size > 0:
            if not quiet:
                print(f"[OK] PDF -> {dst}  （通道 {used}，{dst.stat().st_size // 1024} KB）")
            return dst
        errors.append(f"{name}: 未生成输出文件")

    hint = ("本机未检测到 Microsoft Word，PDF 转换需要 Windows + Office。"
            "替代方案：用 LibreOffice 无头模式 "
            "(`soffice --headless --convert-to pdf <file.docx>`)。"
            if not ensure_word_available() else
            "Word 已安装但转换失败，请检查 Word 是否被其他弹窗/模态框占用。")
    raise PdfExportError("PDF 导出失败：\n  - " + "\n  - ".join(errors) + f"\n提示：{hint}")


def main() -> int:
    ap = argparse.ArgumentParser(description="docx -> PDF 导出")
    ap.add_argument("--input", required=True, help="源 .docx")
    ap.add_argument("--output", default=None, help="目标 .pdf（默认同名）")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args()
    try:
        docx_to_pdf(Path(args.input), args.output, quiet=args.quiet)
    except PdfExportError as e:
        print(f"[error] {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
