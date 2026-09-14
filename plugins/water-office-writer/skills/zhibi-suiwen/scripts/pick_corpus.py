#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 dump_corpus.py 产出的语料里挑出指定文档，拼成一份便于通读的选集。

用法：python pick_corpus.py <语料.txt> <输出.txt> 关键词1 [关键词2 ...]
"""
import sys
from pathlib import Path

src, dst = Path(sys.argv[1]), Path(sys.argv[2])
keys = sys.argv[3:]

raw = src.read_text(encoding="utf-8", errors="replace")
blocks = []
cur_label, buf = None, []
for line in raw.splitlines():
    if line.startswith("#=== file: "):
        if cur_label is not None:
            blocks.append((cur_label, "\n".join(buf)))
        cur_label = line[len("#=== file: "):].rstrip(" =")
        buf = []
    else:
        buf.append(line)
if cur_label is not None:
    blocks.append((cur_label, "\n".join(buf)))

picked = [(l, b) for l, b in blocks if any(k in l for k in keys)]
out = []
for l, b in picked:
    out.append(f"\n{'=' * 70}\n### {l}\n{'=' * 70}\n{b}")
dst.write_text("\n".join(out), encoding="utf-8")
print(f"选中 {len(picked)} 份 -> {dst}；合计 {sum(len(b) for _, b in picked)} 字")
for l, b in picked:
    print(f"  - {l} ({len(b)} 字)")
