#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_style_profile.py - 综合多源样本生成领导风格档案（v3 多源分通道 + format_hints）

用法:
    python build_style_profile.py --intermediate <JSON文件路径模式> --leader <领导姓名> --output <YAML输出>

设计要点：
- v1: 单通道（docx only）
- v2 (2026-09-11): 多源分通道归并，docx/pdf/md 各按权威性合成；字段缺失 → WARNING + source_breakdown 标记
- v3 (2026-09-11): 新增 format_hints 通道
    - md_frontmatter（用户在 MD 顶部 YAML 显式声明的字体/字号）— v2 路径完全感知不到，v3 接进来
    - inferred_from_leader（--inferred-from-leader 触发的跨类型档案注入）— 同上
    - 字段权威性矩阵（v3 扩展）：
        body_font           : docx > pdf > md_frontmatter > inferred
        body_size_pt        : docx > pdf > md_frontmatter > inferred
        line_spacing        : docx > md_frontmatter > inferred       (pdf 无信号)
        first_line_indent_cm: docx > md_frontmatter > inferred > pdf_bool
        page_margins        : docx > md_frontmatter > inferred       (pdf 无信号)
        h1/h2/main_title    : docx > md_frontmatter > inferred       (pdf 仅给 size_pt 不给 family)
    - 非 docx 来源字段：source_breakdown 标注 source: "md_frontmatter"/"inferred"，顶层 warnings 提示"准确度有限"

依赖:
    pip install pyyaml jieba
"""
import argparse
import glob
import json
import os
import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import yaml


# ---------- 多源拆分 ----------
def split_by_type(samples):
    """按 file_type 拆样本"""
    buckets = defaultdict(list)
    for s in samples:
        if s.get("status") != "success":
            continue
        ftype = s.get("file_type", "unknown")
        buckets[ftype].append(s)
    return buckets


# ---------- format_hints 辅助：从 samples 中聚合 ----------
def collect_format_hints(samples):
    """把所有样本的 format_hints 聚合起来。
    返回:
        {
          'md_frontmatter': {key: [val, val, ...]},
          'inferred_from_leader': {key: [val, val, ...]},
        }
    """
    out = {"md_frontmatter": defaultdict(list), "inferred_from_leader": defaultdict(list)}
    for s in samples:
        if s.get("file_type") != "md":
            continue
        hints = s.get("format_hints", {})
        for source_name in ("md_frontmatter", "inferred_from_leader"):
            if source_name in hints:
                applied = hints[source_name].get("applied_fields", {})
                for k, v in applied.items():
                    out[source_name][k].append(v)
    return out


# ---------- 单字段多通道选择器 ----------
def pick_field(channels_with_priority, warnings_list, field_name, default=None):
    """按权威性优先级选择第一个非空值。channels_with_priority 是个 list [(value, source_name), ...]"""
    for value, source in channels_with_priority:
        if value is not None and value != "":
            return value, source
    warnings_list.append(f"{field_name} 全部来源缺失（含 md_frontmatter/inferred）")
    return default, None


# ---------- 格式分析：分通道 + 合成（含 format_hints） ----------
def analyze_format_multi(buckets, hints_pool):
    """分通道统计格式字段，再按权威性合成

    返回 dict: {
        'format': {...},
        'source_breakdown': {...},
        'field_warnings': [...],
    }
    """
    docx = buckets.get("docx", [])
    pdf = buckets.get("pdf", [])
    md = buckets.get("md", [])

    fmt = {}
    breakdown = {}
    warnings = []

    # 辅助：从 docx 样本里统计某字段的 Counter
    def docx_counter(path_fn):
        c = Counter()
        for s in docx:
            v = path_fn(s)
            if v is not None and v != "":
                c[v] += 1
        return c

    # ---------- body_font ----------
    docx_fonts = docx_counter(lambda s: s.get("format_profile", {}).get("fonts", {}).get("body", {}).get("family"))
    pdf_fonts_raw = docx_counter(lambda s: s.get("format_profile", {}).get("fonts", {}).get("body", {}).get("family")) if False else Counter()  # pdf 字段也是 path
    pdf_fonts_raw = Counter()
    for s in pdf:
        f = s.get("format_profile", {}).get("fonts", {}).get("body", {})
        if f and f.get("family"):
            pdf_fonts_raw[f["family"]] += 1

    fm_body_font = hints_pool["md_frontmatter"].get("font_body", [])
    inferred_body_font = hints_pool["inferred_from_leader"].get("font_body", [])

    body_font, body_font_src = pick_field([
        (docx_fonts.most_common(1)[0][0] if docx_fonts else None, "docx"),
        (pdf_fonts_raw.most_common(1)[0][0] if pdf_fonts_raw else None, "pdf"),
        (fm_body_font[0] if fm_body_font else None, "md_frontmatter"),
        (inferred_body_font[0] if inferred_body_font else None, "inferred_from_leader"),
    ], warnings, "body_font")

    fmt["body_font"] = body_font
    fmt["body_font_source"] = body_font_src
    breakdown["body_font"] = {
        "docx": docx_fonts.most_common(3) if docx_fonts else None,
        "pdf": pdf_fonts_raw.most_common(3) if pdf_fonts_raw else None,
        "md_frontmatter": fm_body_font[0] if fm_body_font else None,
        "inferred_from_leader": inferred_body_font[0] if inferred_body_font else None,
        "selected_source": body_font_src,
        "selected_value": body_font,
    }
    if body_font_src in ("md_frontmatter", "inferred_from_leader"):
        warnings.append(f"body_font 来自 {body_font_src}（准确度有限，docx/pdf 缺失时启用）")

    # ---------- body_size_pt ----------
    docx_sizes = Counter()
    for s in docx:
        sz = s.get("format_profile", {}).get("fonts", {}).get("body", {}).get("size_pt")
        if sz:
            docx_sizes[round(sz, 1)] += 1
    pdf_sizes = Counter()
    for s in pdf:
        sz = s.get("format_profile", {}).get("fonts", {}).get("body", {}).get("size_pt")
        if sz:
            pdf_sizes[round(sz, 1)] += 1

    fm_body_size = hints_pool["md_frontmatter"].get("size_body_pt", [])
    inferred_body_size = hints_pool["inferred_from_leader"].get("size_body_pt", [])

    body_size, body_size_src = pick_field([
        (docx_sizes.most_common(1)[0][0] if docx_sizes else None, "docx"),
        (pdf_sizes.most_common(1)[0][0] if pdf_sizes else None, "pdf"),
        (fm_body_size[0] if fm_body_size else None, "md_frontmatter"),
        (inferred_body_size[0] if inferred_body_size else None, "inferred_from_leader"),
    ], warnings, "body_size_pt")

    fmt["body_size_pt"] = body_size
    fmt["body_size_pt_source"] = body_size_src
    breakdown["body_size_pt"] = {
        "docx": dict(docx_sizes.most_common(3)) if docx_sizes else None,
        "pdf": dict(pdf_sizes.most_common(3)) if pdf_sizes else None,
        "md_frontmatter": fm_body_size[0] if fm_body_size else None,
        "inferred_from_leader": inferred_body_size[0] if inferred_body_size else None,
        "selected_source": body_size_src,
        "selected_value": body_size,
    }
    if body_size_src in ("md_frontmatter", "inferred_from_leader"):
        warnings.append(f"body_size_pt 来自 {body_size_src}（准确度有限）")

    # ---------- line_spacing ----------
    docx_lines = Counter()
    for s in docx:
        ls = s.get("format_profile", {}).get("paragraph", {}).get("line_spacing")
        if ls:
            docx_lines[(ls, s["format_profile"]["paragraph"].get("line_spacing_rule"))] += 1
    fm_line_spacing = hints_pool["md_frontmatter"].get("line_spacing_pt", [])
    inferred_line_spacing = hints_pool["inferred_from_leader"].get("line_spacing_pt", [])

    line_spacing, line_spacing_src = pick_field([
        (docx_lines.most_common(1)[0][0][0] if docx_lines else None, "docx"),
        (fm_line_spacing[0] if fm_line_spacing else None, "md_frontmatter"),
        (inferred_line_spacing[0] if inferred_line_spacing else None, "inferred_from_leader"),
    ], warnings, "line_spacing")

    line_spacing_rule = docx_lines.most_common(1)[0][0][1] if docx_lines else None
    fmt["line_spacing"] = line_spacing
    fmt["line_spacing_rule"] = line_spacing_rule
    fmt["line_spacing_source"] = line_spacing_src
    breakdown["line_spacing"] = {
        "docx": docx_lines.most_common(3) if docx_lines else None,
        "pdf": "无可信信号",
        "md_frontmatter": fm_line_spacing[0] if fm_line_spacing else None,
        "inferred_from_leader": inferred_line_spacing[0] if inferred_line_spacing else None,
        "selected_source": line_spacing_src,
    }
    if line_spacing_src in ("md_frontmatter", "inferred_from_leader"):
        warnings.append(f"line_spacing 来自 {line_spacing_src}（pdf 无此信号）")

    # ---------- first_line_indent_cm ----------
    indent_docx_vals = []
    for s in docx:
        v = s.get("format_profile", {}).get("paragraph", {}).get("first_line_indent_cm")
        if v is not None:
            indent_docx_vals.append(v)
    indent_pdf_inferred = [
        s.get("format_profile", {}).get("paragraph", {}).get("first_line_indent_inferred")
        for s in pdf if s.get("format_profile", {}).get("paragraph", {}).get("first_line_indent_inferred") is not None
    ]
    fm_indent = hints_pool["md_frontmatter"].get("first_line_indent_cm", [])
    inferred_indent = hints_pool["inferred_from_leader"].get("first_line_indent_cm", [])

    indent_cm = None
    indent_src = None
    if indent_docx_vals:
        c = Counter(indent_docx_vals)
        indent_cm = c.most_common(1)[0][0]
        indent_src = "docx"
        indent_breakdown = dict(c)
    elif fm_indent:
        indent_cm = fm_indent[0]
        indent_src = "md_frontmatter"
        indent_breakdown = None
        warnings.append("first_line_indent_cm 来自 md_frontmatter（docx 缺失）")
    elif inferred_indent:
        indent_cm = inferred_indent[0]
        indent_src = "inferred_from_leader"
        indent_breakdown = None
        warnings.append("first_line_indent_cm 来自 inferred_from_leader（docx 缺失）")
    elif any(indent_pdf_inferred):
        indent_cm = None  # pdf 只给 bool
        indent_src = "pdf_bool"
        warnings.append("仅 pdf 样本：first_line_indent 仅推断（bool），无 cm 值")

    fmt["first_line_indent_cm"] = indent_cm
    fmt["first_line_indent_source"] = indent_src
    breakdown["first_line_indent_cm"] = {
        "docx": indent_breakdown if indent_docx_vals else None,
        "pdf_bool": indent_pdf_inferred or None,
        "md_frontmatter": fm_indent[0] if fm_indent else None,
        "inferred_from_leader": inferred_indent[0] if inferred_indent else None,
        "selected_source": indent_src,
    }

    # ---------- page_margins ----------
    page_margins = None
    for s in docx:
        p = s.get("format_profile", {}).get("page", {})
        if all(p.get(f"margin_{side}_mm") is not None for side in ("top", "bottom", "left", "right")):
            page_margins = {f"margin_{side}_mm": p[f"margin_{side}_mm"] for side in ("top", "bottom", "left", "right")}
            break

    page_margins_src = None
    if page_margins:
        page_margins_src = "docx"
    else:
        # 从 hints 里找全四个 margin_*_mm
        for src_name in ("md_frontmatter", "inferred_from_leader"):
            side_vals = {}
            all_ok = True
            for side in ("top", "bottom", "left", "right"):
                key = f"page_margin_{side}_mm"
                vals = hints_pool[src_name].get(key, [])
                if vals:
                    side_vals[f"margin_{side}_mm"] = vals[0]
                else:
                    all_ok = False
                    break
            if all_ok:
                page_margins = side_vals
                page_margins_src = src_name
                warnings.append(f"page_margins 来自 {src_name}（docx 缺失）")
                break

    fmt["page_margins_mm"] = page_margins
    fmt["page_margins_source"] = page_margins_src
    breakdown["page_margins_mm"] = {
        "docx": page_margins if page_margins_src == "docx" else None,
        "pdf": "无可信信号",
        "md_frontmatter": {f"margin_{s}_mm": hints_pool["md_frontmatter"].get(f"page_margin_{s}_mm", [None])[0]
                            for s in ("top", "bottom", "left", "right")}
                            if any(hints_pool["md_frontmatter"].get(f"page_margin_{s}_mm") for s in ("top", "bottom", "left", "right"))
                            else None,
        "inferred_from_leader": {f"margin_{s}_mm": hints_pool["inferred_from_leader"].get(f"page_margin_{s}_mm", [None])[0]
                                  for s in ("top", "bottom", "left", "right")}
                                  if any(hints_pool["inferred_from_leader"].get(f"page_margin_{s}_mm") for s in ("top", "bottom", "left", "right"))
                                  else None,
        "selected_source": page_margins_src,
    }
    if pdf or md:
        if not page_margins:
            warnings.append("page_margins 全部来源缺失（含 hints）")

    # ---------- h1/h2/main_title 字体（标题层级字体补全） ----------
    def pick_title_font(role, fm_key, inferred_key):
        """从 docx 拿某 role 的字体，再降级到 hints"""
        c = Counter()
        for s in docx:
            f = s.get("format_profile", {}).get("fonts", {}).get(role, {})
            if f and f.get("family"):
                c[f["family"]] += 1
        fm_vals = hints_pool["md_frontmatter"].get(fm_key, [])
        inf_vals = hints_pool["inferred_from_leader"].get(inferred_key, [])
        val, src = pick_field([
            (c.most_common(1)[0][0] if c else None, "docx"),
            (fm_vals[0] if fm_vals else None, "md_frontmatter"),
            (inf_vals[0] if inf_vals else None, "inferred_from_leader"),
        ], warnings, f"{role}_font")
        return val, src, c, fm_vals, inf_vals

    for role, fm_key, inf_key in [
        ("h1", "font_h1", "font_h1"),
        ("h2", "font_h2", "font_h2"),
        ("main_title", "font_main_title", "font_main_title"),
    ]:
        val, src, c, fm_vals, inf_vals = pick_title_font(role, fm_key, inf_key)
        fmt[f"{role}_font"] = val
        fmt[f"{role}_font_source"] = src
        breakdown[f"{role}_font"] = {
            "docx": c.most_common(3) if c else None,
            "md_frontmatter": fm_vals[0] if fm_vals else None,
            "inferred_from_leader": inf_vals[0] if inf_vals else None,
            "selected_source": src,
        }

    # ---------- headings list（与 v2 一致） ----------
    headings_all = []
    for s in docx:
        for h in s.get("format_profile", {}).get("headings", []):
            headings_all.append({**h, "source": "docx"})
    pdf_heading_count = 0
    for s in pdf:
        for h in s.get("format_profile", {}).get("headings", []):
            headings_all.append({**h, "source": "pdf"})
            pdf_heading_count += 1
    md_heading_levels = []
    for s in md:
        hld = s.get("format_profile", {}).get("paragraph", {}).get("heading_level_distribution", {})
        md_heading_levels.append(hld)
    fmt["headings"] = headings_all[:80]
    fmt["md_heading_level_distribution"] = sum_dicts(md_heading_levels)
    breakdown["headings_count"] = {
        "docx": sum(len(s.get("format_profile", {}).get("headings", [])) for s in docx),
        "pdf": pdf_heading_count,
        "md_levels": fmt["md_heading_level_distribution"],
    }

    return {"format": fmt, "source_breakdown": breakdown, "field_warnings": warnings}


def sum_dicts(dicts):
    out = {}
    for d in dicts:
        if not d:
            continue
        for k, v in d.items():
            try:
                out[k] = out.get(k, 0) + v
            except TypeError:
                pass
    return out


def _sanitize_for_yaml(obj):
    """递归把 tuple 转 list（yaml.dump 遇 tuple 会输出 !!python/tuple tag，safe_load 读不了）"""
    if isinstance(obj, tuple):
        return [_sanitize_for_yaml(x) for x in obj]
    if isinstance(obj, list):
        return [_sanitize_for_yaml(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _sanitize_for_yaml(v) for k, v in obj.items()}
    return obj


# ---------- 语言分析 ----------
def analyze_linguistics_multi(buckets):
    """所有源的文本合并 Counter。每源打 source 标记，方便按需剔除。"""
    try:
        import jieba
        jieba.setLogLevel("ERROR")
    except ImportError:
        print("⚠️ 未安装 jieba，跳过词汇统计")
        return {"linguistic": {}, "warnings": ["jieba 未装"]}

    all_text = []
    source_dist = Counter()
    for ftype, samples in buckets.items():
        for s in samples:
            tp = s.get("text_profile", {})
            paras = tp.get("paragraphs", [])
            if paras:
                all_text.extend(paras)
                source_dist[ftype] += len(paras)
    if not all_text:
        return {"linguistic": {}, "warnings": ["无任何文本样本"], "source_dist": dict(source_dist)}

    full_text = "".join(all_text)

    # 高频四字词
    four_char = re.findall(r"[\u4e00-\u9fff]{4}", full_text)
    four_char_counter = Counter(four_char)
    high_freq = [w for w, c in four_char_counter.most_common(50) if c >= 2][:30]

    # 政策表述
    policy_patterns = [
        r"新发展理念", r"新发展格局", r"高质量发展",
        r"两个维护", r"四个意识", r"四个自信",
        r"五位一体", r"四个全面",
    ]
    political_terms = [p for p in policy_patterns if re.search(p, full_text)]

    # 收尾模式
    closers = []
    for ftype, samples in buckets.items():
        for s in samples:
            ls = s.get("text_profile", {}).get("last_sentences", [])
            if ls:
                closers.extend(ls)
    unique_closers = list(dict.fromkeys(closers))[:5]

    # md 特有的标记统计（如果存在 md）
    md_marks = Counter()
    if "md" in buckets:
        for s in buckets["md"]:
            for b in s.get("text_profile", {}).get("blocks", []) or []:
                for m in (b.get("marks") or []):
                    md_marks[m.get("type", "?")] += 1

    return {
        "linguistic": {
            "high_freq_4char": high_freq,
            "political_terms": political_terms,
            "common_closers": unique_closers,
            "md_mark_distribution": dict(md_marks) if md_marks else None,
        },
        "warnings": [],
        "source_dist": dict(source_dist),
    }


# ---------- 主流程 ----------
def build_profile(samples, leader_name):
    buckets = split_by_type(samples)
    hints_pool = collect_format_hints(samples)
    success_count = sum(len(v) for v in buckets.values())

    fmt_result = analyze_format_multi(buckets, hints_pool)
    lin_result = analyze_linguistics_multi(buckets)

    # hints 通道统计
    fm_count = sum(1 for s in samples if s.get("file_type") == "md" and s.get("format_hints", {}).get("md_frontmatter"))
    inferred_count = sum(1 for s in samples if s.get("file_type") == "md" and s.get("format_hints", {}).get("inferred_from_leader"))

    profile = {
        "meta": {
            "leader_name": leader_name,
            "leader_title": "",
            "organization": "",
            "sample_count": success_count,
            "sample_files": [s["file"] for v in buckets.values() for s in v],
            "sample_breakdown": {k: len(v) for k, v in buckets.items()},
            "format_hints_usage": {
                "md_with_frontmatter": fm_count,
                "md_with_inferred": inferred_count,
            },
            "generated_at": datetime.now().isoformat(),
            "generated_by": "zhibi-suiwen v1.3 (multi-source + format_hints)",
            "version": "v3",
        },
        "format": fmt_result["format"],
        "source_breakdown": fmt_result["source_breakdown"],
        "linguistic": lin_result["linguistic"],
        "text_source_dist": lin_result.get("source_dist", {}),
        "warnings": list(set(fmt_result["field_warnings"] + lin_result.get("warnings", []))),
        "compliance": {
            "follows_gbt_9704": True,
            "deviations": [],
            "notes": "由 build_style_profile.py v3 自动生成。字段权威性：docx > pdf > md_frontmatter > inferred_from_leader。"
                     "非 docx 来源字段在 format.<field>_source 与 source_breakdown.<field>.selected_source 标记，"
                     "并在 warnings 提示。",
        },
    }

    return profile


def load_intermediate(pattern):
    """加载所有中间 JSON 文件"""
    files = sorted(glob.glob(pattern))
    if not files:
        raise FileNotFoundError(f"未找到匹配的中间文件: {pattern}")
    all_data = []
    for f in files:
        with open(f, "r", encoding="utf-8") as fp:
            data = json.load(fp)
            if isinstance(data, list):
                all_data.extend(data)
            else:
                all_data.append(data)
    return all_data


def main():
    parser = argparse.ArgumentParser(description="生成领导风格档案（v3 多源分通道 + format_hints）")
    parser.add_argument("--intermediate", required=True, help="中间 JSON 文件路径模式（如 _intermediate/*.json）")
    parser.add_argument("--leader", required=True, help="领导姓名")
    parser.add_argument("--output", required=True, help="输出 YAML 文件路径")
    args = parser.parse_args()

    pattern = args.intermediate.replace("\\", "/")

    print(f"📂 加载样本：{pattern}")
    samples = load_intermediate(pattern)
    print(f"✅ 共加载 {len(samples)} 份样本")

    print("🔍 分通道分析格式特征（含 md_frontmatter / inferred_from_leader）...")
    print("🔍 分通道分析语言风格...")
    profile = build_profile(samples, args.leader)

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        yaml.dump(_sanitize_for_yaml(profile), f, allow_unicode=True, default_flow_style=False, sort_keys=False)

    print(f"✅ 风格档案已生成：{output_path}")
    print(f"   - 样本数：{profile['meta']['sample_count']}，类型分布：{profile['meta']['sample_breakdown']}")
    print(f"   - hints 使用：frontmatter={profile['meta']['format_hints_usage']['md_with_frontmatter']}, "
          f"inferred={profile['meta']['format_hints_usage']['md_with_inferred']}")
    print(f"   - 正文字体：{profile['format'].get('body_font', 'N/A')}（来源：{profile['format'].get('body_font_source', 'N/A')}）")
    print(f"   - 正文字号：{profile['format'].get('body_size_pt', 'N/A')}（来源：{profile['format'].get('body_size_pt_source', 'N/A')}）")
    print(f"   - 四字词：{len(profile['linguistic'].get('high_freq_4char', []))} 个")
    if profile["warnings"]:
        print(f"⚠️ 字段缺失警告 ({len(profile['warnings'])} 条)：")
        for w in profile["warnings"]:
            print(f"   - {w}")


if __name__ == "__main__":
    main()