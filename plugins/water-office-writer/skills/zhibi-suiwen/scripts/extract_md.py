#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_md.py - 从 .md 文件提取结构与文本（v2 加 frontmatter + 跨类型字体推断）

设计要点：
- v1 (2026-09-11): 手写解析器，识别 ATX heading/列表/引用/code/divider/inline marks
- v2 (2026-09-11):
    - **frontmatter 解析**：MD 顶部 YAML 块 (--- ... ---) 中的字体/字号字段直接覆盖 format_profile 对应字段
    - **跨类型字体推断**：CLI --inferred-from-leader <leader> → 读 skill 内置 profile_<leader>_v*.yaml
      作为 default 注入（与 frontmatter 同等优先级；都打 source 标记，绝不静默）
    - **format_hints 通道**：所有非直接来源（md_frontmatter / inferred_from_leader）的字段单独保留在
      result["format_hints"] 字段，供 build_style_profile 多源合成时按权威性矩阵读取

依赖:
    必需: pip install pyyaml
"""
import argparse
import glob
import json
import os
import re
from collections import Counter
from pathlib import Path

import yaml


# ---------- 正则定义 ----------
RE_ATX_HEADING = re.compile(r"^(#{1,6})\s+(.+?)\s*#*\s*$")
RE_SETH_HEADING = re.compile(r"^([=-]{2,})\s*$")  # 仅在段落首行下一行匹配
RE_HR = re.compile(r"^[\s]*(?:[\-\*_]){3,}[\s]*$")
RE_FENCE = re.compile(r"^```(\w*)")
RE_OL_ITEM = re.compile(r"^\s*\d+[.)、]\s+")
RE_UL_ITEM = re.compile(r"^\s*[-*+]\s+")
RE_QUOTE = re.compile(r"^\s*(?:>\s*)+")
RE_INLINE_CODE = re.compile(r"`([^`]+)`")
RE_BOLD = re.compile(r"(\*\*|__)(.+?)\1")
RE_ITALIC = re.compile(r"(\*|_)(.+?)\1")

# 行内 mark 提取顺序：先代码，再强调
# italic 用 (?<!\*)/(?!\*) 防止与 bold 重叠：`**xxx**` 内的两端 * 不会被识别为 italic
INLINE_PATTERNS = [
    ("code", RE_INLINE_CODE),
    ("bold", RE_BOLD),
    ("italic", re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|(?<![_\w])_([^_\s]+)_(?![_\w])")),
]

# frontmatter 与 format_profile 字段映射
# 键：frontmatter 字段名 → 值：(format_profile 内路径, 类型转换函数)
# 路径以 result["format_profile"] 为根（不含 "format_profile." 前缀）
FONT_FIELD_MAP = {
    "font_body":         ("fonts.body.family",          str),
    "size_body_pt":      ("fonts.body.size_pt",         float),
    "font_h1":           ("fonts.h1.family",            str),
    "size_h1_pt":        ("fonts.h1.size_pt",           float),
    "font_h2":           ("fonts.h2.family",            str),
    "size_h2_pt":        ("fonts.h2.size_pt",           float),
    "font_main_title":   ("fonts.main_title.family",    str),
    "size_main_title_pt":("fonts.main_title.size_pt",   float),
    "first_line_indent_cm": ("paragraph.first_line_indent_cm", float),
    "line_spacing_pt":   ("paragraph.line_spacing",      float),
    "page_margin_top_mm":    ("page.margin_top_mm",      float),
    "page_margin_bottom_mm": ("page.margin_bottom_mm",   float),
    "page_margin_left_mm":   ("page.margin_left_mm",     float),
    "page_margin_right_mm":  ("page.margin_right_mm",    float),
}


def extract_marks(text):
    """提取行内强调/code 的位置和类型。返回 [{type, text}, ...]"""
    marks = []
    for mtype, pat in INLINE_PATTERNS:
        for m in pat.finditer(text):
            # 对 RE_BOLD/RE_ITALIC（有 2 个捕获组）：m.group(2) 是内容
            # 对 RE_INLINE_CODE（1 个捕获组）：m.group(1) 是内容
            inner = m.group(2) if m.lastindex and m.lastindex >= 2 else m.group(1)
            marks.append({"type": mtype, "text": inner})
    return marks


def parse_markdown_lines(lines):
    """逐行解析，返回 (blocks, headings)

    blocks: list[dict]，每个元素至少含 block_type/text
    """
    blocks = []
    headings = []

    i = 0
    n = len(lines)
    cur_paragraph = []  # 当前累积的段落行
    cur_list_items = []  # 当前累积的列表
    cur_list_kind = None  # 'ul' / 'ol' / None
    cur_quote_lines = []

    def flush_paragraph():
        nonlocal cur_paragraph
        if cur_paragraph:
            text = " ".join(cur_paragraph).strip()
            if text:
                marks = extract_marks(text)
                blocks.append({
                    "block_type": "paragraph",
                    "text": text,
                    "marks": marks,
                })
            cur_paragraph = []

    def flush_list():
        nonlocal cur_list_items, cur_list_kind
        if cur_list_items:
            kind = cur_list_kind or "ul"
            items_text = []
            for itm in cur_list_items:
                items_text.append(itm)
            blocks.append({
                "block_type": "list",
                "list_kind": kind,
                "items": items_text,
                "marks": [m for it in items_text for m in extract_marks(it)],
            })
            cur_list_items = []
            cur_list_kind = None

    def flush_quote():
        nonlocal cur_quote_lines
        if cur_quote_lines:
            text = " ".join(cur_quote_lines).strip()
            if text:
                blocks.append({
                    "block_type": "quote",
                    "text": text,
                    "marks": extract_marks(text),
                })
            cur_quote_lines = []

    while i < n:
        raw = lines[i]
        line = raw.rstrip("\n")

        # 空行：结束任何块
        if not line.strip():
            flush_paragraph()
            flush_list()
            flush_quote()
            i += 1
            continue

        # 代码块围栏
        if RE_FENCE.match(line.strip()):
            flush_paragraph()
            flush_list()
            flush_quote()
            # 找结束围栏
            fence_lang = RE_FENCE.match(line.strip()).group(1)
            i += 1
            code_lines = []
            while i < n and not RE_FENCE.match(lines[i].rstrip("\n").strip()):
                code_lines.append(lines[i].rstrip("\n"))
                i += 1
            blocks.append({
                "block_type": "code",
                "language": fence_lang or None,
                "text": "\n".join(code_lines),
            })
            i += 1  # 跳过结束围栏
            continue

        # 水平线
        if RE_HR.match(line):
            flush_paragraph()
            flush_list()
            flush_quote()
            blocks.append({"block_type": "divider", "text": line.strip()})
            i += 1
            continue

        # ATX 标题
        m_atx = RE_ATX_HEADING.match(line)
        if m_atx:
            flush_paragraph()
            flush_list()
            flush_quote()
            level = len(m_atx.group(1))
            title_text = m_atx.group(2).strip()
            # 去掉行内 mark 仅用于展示
            marks = extract_marks(title_text)
            headings.append({"level": level, "text": title_text, "size_pt": None, "fontname": None})
            blocks.append({
                "block_type": "heading",
                "level": level,
                "text": title_text,
                "marks": marks,
            })
            i += 1
            continue

        # 列表项
        if RE_UL_ITEM.match(line) or RE_OL_ITEM.match(line):
            flush_paragraph()
            flush_quote()
            kind = "ol" if RE_OL_ITEM.match(line) else "ul"
            content = (RE_UL_ITEM.match(line) and RE_UL_ITEM.sub("", line, count=1)) or (
                RE_OL_ITEM.match(line) and RE_OL_ITEM.sub("", line, count=1)
            )
            if cur_list_kind and cur_list_kind != kind:
                flush_list()
            cur_list_kind = kind
            cur_list_items.append(content.strip())
            i += 1
            continue

        # 引用
        m_quote = RE_QUOTE.match(line)
        if m_quote:
            flush_paragraph()
            flush_list()
            content = RE_QUOTE.sub("", line)
            cur_quote_lines.append(content.strip())
            i += 1
            continue

        # 普通段落行
        flush_list()
        flush_quote()
        cur_paragraph.append(line.strip())
        i += 1

    flush_paragraph()
    flush_list()
    flush_quote()

    return blocks, headings


def parse_frontmatter(content):
    """解析 MD 顶部 YAML frontmatter（--- ... ---）。
    返回 (frontmatter_dict or None, body_text_without_frontmatter)
    """
    if not content.startswith("---"):
        return None, content
    lines = content.split("\n")
    if not lines or lines[0].rstrip() != "---":
        return None, content
    # 找第二个 ---
    end = None
    for i in range(1, len(lines)):
        if lines[i].rstrip() == "---":
            end = i
            break
    if end is None:
        return None, content
    yaml_text = "\n".join(lines[1:end])
    body_text = "\n".join(lines[end + 1:])
    try:
        data = yaml.safe_load(yaml_text)
        if not isinstance(data, dict):
            return None, content
        return data, body_text
    except yaml.YAMLError:
        return None, content


def _set_by_dotted_path(d, path, value):
    """按 "a.b.c" 路径设置嵌套 dict 的值"""
    parts = path.split(".")
    cur = d
    for p in parts[:-1]:
        if p not in cur or not isinstance(cur[p], dict):
            cur[p] = {}
        cur = cur[p]
    cur[parts[-1]] = value


def _get_by_dotted_path(d, path):
    parts = path.split(".")
    cur = d
    for p in parts:
        if not isinstance(cur, dict) or p not in cur:
            return None
        cur = cur[p]
    return cur


def apply_format_hints(result, hints_dict, source):
    """把 hints 字典按 FONT_FIELD_MAP 覆盖到 result 的 format_profile 对应字段。
    同时记入 result['format_hints'] 通道，标注每个字段来源。
    source: "md_frontmatter" | "inferred_from_leader"
    """
    if not hints_dict:
        return
    applied = {}
    for fm_key, (target_path, caster) in FONT_FIELD_MAP.items():
        if fm_key not in hints_dict:
            continue
        raw = hints_dict[fm_key]
        try:
            value = caster(raw)
        except (ValueError, TypeError):
            continue
        _set_by_dotted_path(result["format_profile"], target_path, value)
        applied[fm_key] = value
    if applied:
        # 记录整批 hints（含 source_intent 等元字段）
        result.setdefault("format_hints", {})[source] = {
            "applied_fields": applied,
            "raw": {k: v for k, v in hints_dict.items() if k in FONT_FIELD_MAP or k == "source_intent"},
        }


def load_leader_profile(leader_name, skill_root):
    """从 skill 内置 assets/ 目录读取 profile_<leader>_v*.yaml。
    优先项目级 skill，回退到脚本同目录 assets。
    返回扁平化的 hints dict（与 frontmatter 同字段集），找不到返回 None。
    """
    # 路径候选：项目级 skill assets
    candidates = []
    project_root = Path(skill_root) if skill_root else None
    if project_root:
        candidates.append(project_root / "assets")
    # 脚本同目录 assets（分发包 fallback）
    candidates.append(Path(__file__).parent.parent / "assets")

    for assets_dir in candidates:
        if not assets_dir.exists():
            continue
        # profile_<leader>_v*.yaml，支持 "龚利民"、"方琳"、"龚董"、"方总" 别名
        leader_aliases = _LEADER_ALIASES.get(leader_name, [leader_name])
        for alias in leader_aliases:
            matches = sorted(assets_dir.glob(f"profile_{alias}_v*.yaml"))
            if matches:
                with open(matches[0], "r", encoding="utf-8") as f:
                    data = yaml.safe_load(f)
                return _profile_to_hints(data, matches[0].name)
    return None


def _profile_to_hints(profile_data, profile_filename):
    """把 profile_<leader>_v*.yaml 转成与 frontmatter 同结构 hints dict。"""
    fmt = profile_data.get("format", {})
    fonts = fmt.get("fonts", {})
    paragraph = fmt.get("paragraph", {})
    page = fmt.get("page", {})

    hints = {}

    # 字体
    body_font = fonts.get("body", {})
    if body_font.get("family"):
        hints["font_body"] = body_font["family"]
    if body_font.get("size_pt"):
        hints["size_body_pt"] = body_font["size_pt"]

    h1 = fonts.get("h1", {})
    if h1.get("family"):
        hints["font_h1"] = h1["family"]
    if h1.get("size_pt"):
        hints["size_h1_pt"] = h1["size_pt"]

    h2 = fonts.get("h2", {})
    if h2.get("family"):
        hints["font_h2"] = h2["family"]
    if h2.get("size_pt"):
        hints["size_h2_pt"] = h2["size_pt"]

    main_title = fonts.get("main_title", {})
    if main_title.get("family"):
        hints["font_main_title"] = main_title["family"]
    if main_title.get("size_pt"):
        hints["size_main_title_pt"] = main_title["size_pt"]

    # 段落
    if paragraph.get("first_line_indent_cm"):
        hints["first_line_indent_cm"] = paragraph["first_line_indent_cm"]
    if paragraph.get("line_spacing_value"):
        hints["line_spacing_pt"] = paragraph["line_spacing_value"]

    # 页面
    margins = page.get("margins_mm", {})
    for side in ("top", "bottom", "left", "right"):
        key = f"page_margin_{side}_mm"
        if margins.get(side) is not None:
            hints[key] = margins[side]

    # 元信息
    meta = profile_data.get("meta", {})
    hints["source_intent"] = f"推断自内置档案 {profile_filename}（leader={meta.get('leader_name', '?')}）"

    return hints


# 领导别名映射（与 md_to_gongwen_docx.py 保持一致）
_LEADER_ALIASES = {
    "龚利民": ["龚利民", "龚董"],
    "龚董": ["龚董", "龚利民"],
    "方琳": ["方琳", "方总"],
    "方总": ["方总", "方琳"],
    "总裁": ["方琳", "方总"],
    "董事长": ["龚利民", "龚董"],
}


def extract_one_md(file_path, inferred_from_leader=None, skill_root=None):
    """处理单个 MD 文件

    inferred_from_leader: 提供时（如 "龚利民"），先读内置 profile 注入 hints；
                          仍允许 frontmatter 覆盖（同字段 frontmatter 优先于 inferred）
    """
    result = {
        "file": os.path.basename(file_path),
        "file_type": "md",
        "format_profile": {
            "fonts": {
                "body": {"family": None, "size_pt": None, "color": None},
            },
            "paragraph": {},
            "page": {},
            "headings": [],
        },
        "text_profile": {},
        "format_hints": {},
        "warnings": [],
        "meta": {
            "has_frontmatter": False,
            "frontmatter_keys": [],
            "inferred_from_leader": inferred_from_leader or None,
        },
    }

    try:
        with open(str(file_path), "r", encoding="utf-8") as f:
            content = f.read()

        # ---- frontmatter 解析（先做，得到 hints 优先） ----
        fm_data, body_content = parse_frontmatter(content)
        if fm_data:
            result["meta"]["has_frontmatter"] = True
            result["meta"]["frontmatter_keys"] = sorted(fm_data.keys())
            apply_format_hints(result, fm_data, source="md_frontmatter")
            if not fm_data.get("source_intent"):
                # 自动补一段 source_intent 写入 hints
                pass

        # ---- inferred from leader（次优级，frontmatter 有字段时跳过同字段） ----
        if inferred_from_leader:
            leader_hints = load_leader_profile(inferred_from_leader, skill_root)
            if leader_hints:
                # 如果已有 frontmatter hints，从中剔除非 frontmatter 独有字段
                if "md_frontmatter" in result["format_hints"]:
                    # frontmatter 字段优先 → leader_hints 中与 frontmatter 同字段的跳过
                    fm_keys = set(result["format_hints"]["md_frontmatter"]["applied_fields"].keys())
                    filtered_leader_hints = {k: v for k, v in leader_hints.items() if k not in fm_keys}
                    if filtered_leader_hints:
                        apply_format_hints(result, filtered_leader_hints, source="inferred_from_leader")
                else:
                    apply_format_hints(result, leader_hints, source="inferred_from_leader")
            else:
                result["warnings"].append(f"--inferred-from-leader={inferred_from_leader} 未找到对应内置档案")

        # ---- 解析正文 ----
        lines = body_content.split("\n")
        blocks, headings = parse_markdown_lines(lines)

        # ---- 文本档案 ----
        text_blocks = [b for b in blocks if b["block_type"] in ("paragraph", "heading", "list", "quote")]
        full_text_parts = []
        paragraph_count = 0
        for b in blocks:
            if b["block_type"] == "paragraph":
                full_text_parts.append(b["text"])
                paragraph_count += 1
            elif b["block_type"] == "list":
                full_text_parts.extend(b.get("items", []))
                paragraph_count += 1
            elif b["block_type"] == "quote":
                full_text_parts.append(b["text"])
                paragraph_count += 1
            elif b["block_type"] == "heading":
                full_text_parts.append(b["text"])
        full_text = "\n".join(full_text_parts)

        first_sentences = []
        last_sentences = []
        for b in blocks:
            if b["block_type"] == "paragraph" and b["text"]:
                first_sentences.append(b["text"][:50])
                last_sentences.append(b["text"][-50:])

        result["text_profile"] = {
            "total_chars": len(full_text),
            "paragraph_count": paragraph_count,
            "paragraphs": full_text_parts,
            "first_sentences": first_sentences[:10],
            "last_sentences": last_sentences[-5:],
            "blocks": blocks,
        }

        # ---- 格式档案 ----
        # body 字段默认值
        if result["format_profile"]["fonts"]["body"]["family"] is None:
            result["format_profile"]["fonts"]["body"] = {
                "family": None, "size_pt": None, "color": None,
            }
        else:
            # 已有 hints 注入 → 标记来源
            src = "md_frontmatter" if "md_frontmatter" in result["format_hints"] else (
                "inferred_from_leader" if "inferred_from_leader" in result["format_hints"] else None)
            if src:
                result["format_profile"]["fonts"]["body"]["_source"] = src

        # 段落
        level_counter = Counter(h["level"] for h in headings)
        para = result["format_profile"]["paragraph"]
        para.setdefault("first_line_indent_inferred", None)
        para.setdefault("line_spacing", None)
        para.setdefault("line_spacing_rule", None)
        para.setdefault("space_before_pt", None)
        para.setdefault("space_after_pt", None)
        para["heading_level_distribution"] = dict(level_counter)

        # 页面
        page = result["format_profile"]["page"]
        page.setdefault("paper_width_mm", None)
        page.setdefault("paper_height_mm", None)
        for side in ("top", "bottom", "left", "right"):
            page.setdefault(f"margin_{side}_mm", None)

        # 兜底警告
        body = result["format_profile"]["fonts"]["body"]
        if body["family"] is None and body["size_pt"] is None:
            result["warnings"].append(
                "MD 字体字号完全缺失。可选修复：(1) 在 MD 顶部加 YAML frontmatter 声明字体；"
                "(2) 跑 extract_md.py 时加 --inferred-from-leader <leader>"
            )
        else:
            # 有部分 hints，列出哪些字段被覆盖
            sources = result["format_hints"]
            if sources:
                for src_name, src_data in sources.items():
                    applied = src_data["applied_fields"]
                    result["warnings"].append(
                        f"MD 格式字段来自 {src_name}（{len(applied)} 个）：{', '.join(sorted(applied.keys()))}"
                    )

        # 标题
        result["format_profile"]["headings"] = headings[:50]
        result["format_profile"]["headings_count"] = len(headings)
        result["block_stats"] = dict(Counter(b["block_type"] for b in blocks))

        result["status"] = "success"
    except Exception as e:
        result["status"] = "failed"
        result["error"] = str(e)

    return result


def main():
    parser = argparse.ArgumentParser(
        description="从 Markdown 提取结构与文本（v2：frontmatter + 跨类型字体推断）"
    )
    parser.add_argument("--input", required=True, help="输入文件或目录")
    parser.add_argument("--output", required=True, help="输出 JSON 文件")
    parser.add_argument("--inferred-from-leader", default=None,
                        help="当 MD 无 frontmatter 时，从 skill 内置 profile_<leader>_v*.yaml 注入字体默认值"
                             "（如：龚利民 / 龚董 / 方琳 / 方总 / 总裁 / 董事长）")
    args = parser.parse_args()

    input_path = Path(args.input)
    if input_path.is_dir():
        files = sorted(input_path.glob("**/*.md"))
        # 排除 .workbuddy 与 .git 内部
        files = [f for f in files if ".workbuddy" not in str(f) and ".git" not in str(f)]
    else:
        files = [input_path]

    if not files:
        print(f"⚠️ 未找到 MD 文件：{args.input}")
        return

    results = [extract_one_md(f, inferred_from_leader=args.inferred_from_leader) for f in files]

    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    success_count = sum(1 for r in results if r["status"] == "success")
    fm_count = sum(1 for r in results if r.get("meta", {}).get("has_frontmatter"))
    inferred_count = sum(1 for r in results if r.get("format_hints", {}).get("inferred_from_leader"))
    print(f"✅ 处理完成：{success_count}/{len(results)} 成功，输出至 {args.output}")
    if fm_count:
        print(f"📋 含 frontmatter：{fm_count} 个")
    if inferred_count:
        print(f"🔗 应用 leader 推断：{inferred_count} 个")
    warn_count = sum(len(r.get("warnings", [])) for r in results)
    if warn_count:
        print(f"⚠️ 共 {warn_count} 条警告（见各文件 warnings 字段）")


if __name__ == "__main__":
    main()