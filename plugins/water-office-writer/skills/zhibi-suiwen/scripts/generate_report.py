#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_report.py — 基于风格档案生成报告【骨架源稿 + 落版】

设计变更（v2）:
    旧版把"档案解析 + 渲染"各写一遍，且字段名与资产档案不符
    （读 format.body_size_pt，实际是 format.fonts.body.size_pt），
    导致字号/边距/纸张静默回退、A4 变成 Letter、缩进不随字号联动。

    v2 改为「单一落版真源」：
        profile YAML → 骨架 Markdown（结构/称呼/落款按档案生成）
                     → md_to_gongwen_docx.build() 落版
    版式参数只在 md_to_gongwen_docx 里定义一次，本脚本不再复制渲染逻辑。

⚠️ 正文内容仍为骨架占位：真正撰写由模型完成 —— 模型应参照本脚本
   生成的 Markdown 结构，用风格档案的语汇把占位段落改写成正式文稿，
   再调用 md_to_gongwen_docx.py 落版。

用法:
    python generate_report.py --profile assets/profile_龚利民_总结_v1.yaml \
        --topic "2026年三季度生产经营分析会讲话" \
        --outline "一、成绩|二、形势|三、任务|四、号召" \
        --length 4000 --output outputs/2026Q3_讲话_龚利民.docx

依赖:
    pip install python-docx pyyaml
"""
import argparse
import re
from datetime import datetime
from pathlib import Path

import yaml

import md_to_gongwen_docx as m2d

# 档案 structure.typical_outline_* 的默认骨架（档案缺失时兜底）
DEFAULT_OUTLINE = [
    "一、回顾过去工作",
    "二、分析当前形势",
    "三、部署下一步任务",
    "四、号召与要求",
]


def _dig(d, *keys, default=None):
    cur = d
    for k in keys:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur if cur is not None else default


def load_profile_dict(path: Path) -> dict:
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if "format" not in data and "meta" not in data:
        raise SystemExit(f"[error] 不是有效的风格档案: {path}")
    return data


def pick_outline(profile: dict, kind: str | None = None) -> list[str]:
    """从档案 linguistic.structure 里取结构范式；取不到则用默认四段式"""
    st = _dig(profile, "linguistic", "structure", default={}) or {}
    for key in ([f"typical_outline_{kind}"] if kind else []) + \
               ["typical_outline_annual", "typical_outline_halfyear"]:
        v = st.get(key)
        if isinstance(v, list) and v:
            items = []
            for s in v:
                s = str(s)
                # 「一、XXX —— 说明」→ 只保留「一、XXX」
                s = re.split(r"\s*——\s*", s)[0]
                s = re.split(r"\s*[（(]示例", s)[0]
                items.append(s.strip())
            return [i for i in items if i]
    return list(DEFAULT_OUTLINE)


def build_skeleton(profile: dict, topic: str, outline: list[str],
                   length: int, leader: str) -> str:
    """按档案结构生成 Markdown 骨架（供模型改写为正式文稿）"""
    sig = profile.get("signature", {}) or {}
    name_format = sig.get("name_format") or _dig(profile, "meta", "leader_name",
                                                 default=leader) or "（署名待确认）"
    date_fmt = sig.get("date_format", "")
    today = datetime.now()
    date_line = (f"（{today.year}年{today.month}月{today.day}日）"
                 if "（" in date_fmt or "(" in date_fmt
                 else f"{today.year}年{today.month}月{today.day}日")

    per_section = max(200, length // max(1, len(outline)))
    lines = [
        f"# {topic}",
        "",
        f"> 【骨架源稿】本文件由 generate_report.py 依档案结构生成，正文为占位，"
        f"请由模型按 profile 语汇改写后调用 md_to_gongwen_docx.py 落版。",
        f"> 目标字数 {length}，平均每节约 {per_section} 字。",
        f"> 版式档案：{_dig(profile, 'meta', 'leader_name', default=leader)}"
        f"（{_dig(profile, 'meta', 'leader_title', default='')}）",
        "",
    ]
    opener = _dig(profile, "linguistic", "sentence_patterns", "opener", default=[]) or []
    if opener:
        lines.append(str(opener[0]))
        lines.append("")
        lines.append(f"今天召开这次会议，主要任务是研究部署{topic}有关工作。")
        lines.append("")

    for item in outline:
        title = item
        m = re.match(r"^([一二三四五六七八九十]+、)\s*(.*)$", item)
        if m:
            lines.append(f"### {item}")
        else:
            lines.append(f"### {item}")
        lines.append("")
        lines.append(f"（一）关于{title}的第一项举措。[占位 {per_section // 3} 字："
                     f"引档案语汇与真实数据，说明目标、措施、责任单位。]")
        lines.append("")
        lines.append(f"（二）关于{title}的第二项举措。[占位 {per_section // 3} 字："
                     f"补充数据支撑与时间节点。]")
        lines.append("")
        lines.append(f"（三）关于{title}的第三项举措。[占位 {per_section // 3} 字："
                     f"表述要求与考核安排。]")
        lines.append("")

    closer = _dig(profile, "linguistic", "sentence_patterns", "closer", default=[]) or []
    if closer:
        lines.append(str(closer[0]))
        lines.append("")
    lines += ["", f"**{name_format}**", f"**{date_line}**", ""]
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="基于风格档案生成报告（骨架源稿 + 落版）")
    ap.add_argument("--profile", required=True, help="风格档案 YAML 路径")
    ap.add_argument("--topic", required=True, help="报告主题")
    ap.add_argument("--outline", default="", help="大纲（用 | 分隔；留空则取档案结构范式）")
    ap.add_argument("--length", type=int, default=3000, help="目标字数")
    ap.add_argument("--output", required=True, help="输出 .docx 路径")
    ap.add_argument("--md-out", default=None, help="骨架 Markdown 输出路径（默认与 docx 同名）")
    ap.add_argument("--outline-kind", choices=["annual", "halfyear"], default=None,
                    help="取档案里哪套结构范式")
    ap.add_argument("--clause-scope", choices=["lead", "whole"], default="lead")
    ap.add_argument("--no-docx", action="store_true", help="只出 Markdown 骨架，不落版")
    ap.add_argument("--pdf", action="store_true",
                    help="同时导出 PDF（需 Windows + Microsoft Word）")
    ap.add_argument("--pdf-out", default=None, help="PDF 输出路径（默认与 docx 同名）")
    args = ap.parse_args()

    prof_path = Path(args.profile)
    if not prof_path.is_file():
        raise SystemExit(f"[error] 找不到风格档案: {prof_path}")
    profile = load_profile_dict(prof_path)

    leader = _dig(profile, "meta", "leader_name", default="")
    if args.outline:
        outline = [s.strip() for s in args.outline.split("|") if s.strip()]
    else:
        outline = pick_outline(profile, args.outline_kind)

    md_text = build_skeleton(profile, args.topic, outline, args.length, leader)

    out_docx = Path(args.output)
    md_path = Path(args.md_out) if args.md_out else out_docx.with_suffix(".md")
    md_path.parent.mkdir(parents=True, exist_ok=True)
    md_path.write_text(md_text, encoding="utf-8")
    print(f"[OK] 骨架 Markdown -> {md_path}  （{len(md_text)} 字，{len(outline)} 节）")

    if args.no_docx:
        return
    out_docx.parent.mkdir(parents=True, exist_ok=True)
    m2d.build(md_path, out_docx, leader,
              profile_path=str(prof_path),
              clause_scope=args.clause_scope,
              assets_dir=prof_path.parent,
              pdf=args.pdf, pdf_path=args.pdf_out)


if __name__ == "__main__":
    main()
