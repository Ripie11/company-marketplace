#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dump_corpus.py —— 把知识库样本全量转成纯文本语料，供金句挖掘（mine_phrases.py / 人工通读）使用。

用法:
    python dump_corpus.py --kb-root "E:/深圳水务集团/2.2-数字员工项目/知识库" \
                          --out-dir "E:/深圳水务集团/2.2-数字员工项目/_backtest/corpus"

说明:
    - docx: python-docx 逐段提取（含表格文本）
    - pdf : pdfplumber 文本层；文本层为空则报 WARN 并跳过（扫描件走 extract_pdf.py --ocr）
    - doc : 老二进制格式，python-docx 不支持，尝试 pywin32(Word COM) 转换；失败则记 SKIP
    - pptx: 只取文本框，忽略图片（超大 pptx 可选 --skip-pptx）

产出:
    corpus_<领导>.txt  每份样本以 "#=== file: <路径> ===" 分隔
    dump_report.json   每份文件的提取状态（ok / skip / warn + 字数）
"""
import argparse
import json
import sys
from pathlib import Path

TEXT_EXT = {".docx", ".pdf", ".doc", ".pptx"}


def read_docx(path: Path) -> str:
    from docx import Document

    doc = Document(str(path))
    out = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            out.append(t)
    for tb in doc.tables:
        for row in tb.rows:
            cells = [c.text.strip().replace("\n", " ") for c in row.cells]
            if any(cells):
                out.append(" | ".join(cells))
    return "\n".join(out)


def read_pdf(path: Path) -> str:
    import pdfplumber

    out = []
    with pdfplumber.open(str(path)) as pdf:
        for i, page in enumerate(pdf.pages):
            t = page.extract_text() or ""
            if t.strip():
                out.append(t)
    return "\n".join(out)


def read_pptx(path: Path) -> str:
    from pptx import Presentation

    prs = Presentation(str(path))
    out = []
    for idx, slide in enumerate(prs.slides, 1):
        buf = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    t = "".join(r.text for r in p.runs).strip()
                    if t:
                        buf.append(t)
        if buf:
            out.append(f"[slide {idx}] " + " / ".join(buf))
    return "\n".join(out)


def read_doc(path: Path) -> str:
    """老 .doc：尝试 Word COM 另存为临时 docx 再读。"""
    try:
        import win32com.client  # type: ignore
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(f"pywin32 不可用：{e}")

    import tempfile

    tmp = Path(tempfile.gettempdir()) / (path.stem + "_conv.docx")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        d = word.Documents.Open(str(path), ReadOnly=True)
        d.SaveAs(str(tmp), FileFormat=16)  # 16 = wdFormatXMLDocument
        d.Close(False)
    finally:
        word.Quit()
    txt = read_docx(tmp)
    try:
        tmp.unlink()
    except OSError:
        pass
    return txt


def leader_of(path: Path, kb_root: Path) -> str:
    rel = path.relative_to(kb_root)
    top = rel.parts[0]
    if "龚利民" in top:
        return "龚利民"
    if "方琳" in top:
        return "方琳"
    return "通用"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--kb-root", required=True)
    ap.add_argument("--out-dir", required=True)
    ap.add_argument("--skip-pptx", action="store_true", help="跳过 pptx（超大文件建议开启）")
    args = ap.parse_args()

    kb_root = Path(args.kb_root).resolve()
    out_dir = Path(args.out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    buckets: dict[str, list[str]] = {}
    report = []

    files = sorted(
        p for p in kb_root.rglob("*")
        if p.is_file() and p.suffix.lower() in TEXT_EXT and "_build" not in p.parts
    )
    for f in files:
        ext = f.suffix.lower()
        if args.skip_pptx and ext == ".pptx":
            report.append({"file": str(f), "status": "skip", "reason": "--skip-pptx"})
            continue
        try:
            if ext == ".docx":
                txt = read_docx(f)
            elif ext == ".pdf":
                txt = read_pdf(f)
            elif ext == ".pptx":
                txt = read_pptx(f)
            else:
                txt = read_doc(f)
        except Exception as e:  # noqa: BLE001
            report.append({"file": str(f), "status": "skip", "reason": f"{type(e).__name__}: {e}"})
            print(f"[SKIP] {f.name} -> {e}", file=sys.stderr)
            continue

        n = len(txt.strip())
        if n < 200:
            report.append({"file": str(f), "status": "warn", "chars": n,
                           "reason": "文本层过少（可能为扫描件，需 OCR）"})
            print(f"[WARN] {f.name} -> 仅 {n} 字（疑似扫描件）", file=sys.stderr)
            continue

        lg = leader_of(f, kb_root)
        buckets.setdefault(lg, []).append(f"#=== file: {f.relative_to(kb_root).as_posix()} ===\n{txt}")
        report.append({"file": str(f), "status": "ok", "chars": n, "leader": lg})
        print(f"[OK]   {f.name} ({n} 字) -> {lg}")

    for lg, chunks in buckets.items():
        out = out_dir / f"corpus_{lg}.txt"
        out.write_text("\n\n".join(chunks), encoding="utf-8")
        total = sum(len(c) for c in chunks)
        print(f"\n写出 {out} —— {len(chunks)} 份样本，合计 {total} 字")

    (out_dir / "dump_report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
