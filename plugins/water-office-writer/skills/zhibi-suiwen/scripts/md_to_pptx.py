#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md_to_pptx.py - 把 markdown 源稿按领导风格生成 PPTX

用法:
    python md_to_pptx.py --input <md> --output <pptx> [--leader 龚利民]
                        [--profile <profile.yaml>] [--ratio 16:9]

依赖:
    pip install python-pptx pyyaml

设计原则:
    - 一页 = 一个「条」（成效/不足）或一个「要求」/「章节封面」
    - 标题/正文/条款字体严格沿用 leader profile
    - 配色：蓝色主调（#1F5FA8）+ 国企红点缀（#C8102E）
    - 16:9 宽屏，便于会议投屏
"""
import argparse
import re
import sys
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.util import Pt, Inches, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn


# ============================================================
# 默认风格（龚董 profile_龚利民_总结_v1.yaml 同源）
# ============================================================
DEFAULT_PROFILE = {
    "fonts": {
        "title_main":   "方正小标宋简体",
        "title_sub":    "黑体",
        "body":         "仿宋_GB2312",
        "clause":       "楷体",
        "number":       "黑体",
    },
    "sizes_pt": {
        "title_main":   40,    # 封面主标题
        "title_sub":    24,    # 封面副标题
        "page_title":   30,    # 章节封面
        "section_h1":   28,    # 页内大标题
        "body":         18,
        "clause":       20,
        "footer":       10,
    },
    "colors": {
        "primary":      "1F5FA8",   # 水务蓝
        "accent":       "C8102E",   # 国企红
        "dark":         "1A3A66",   # 深蓝
        "muted":        "6B7280",   # 灰
        "light":        "E8F0FA",   # 浅蓝底
        "text":         "1A1A1A",
        "white":        "FFFFFF",
    },
    "footer_text":   "深圳环境水务集团 · 2026年上半年工作总结推进会议",
    "page_w_in":     13.333,
    "page_h_in":     7.5,
    "ratio":         "16:9",
}


# ============================================================
# 工具函数
# ============================================================
def load_profile(path):
    if not path:
        return DEFAULT_PROFILE
    p = Path(path)
    if not p.exists():
        return DEFAULT_PROFILE
    with open(p, encoding="utf-8") as f:
        user = yaml.safe_load(f) or {}
    # 浅合并（不强制覆盖字体，但接受 sizes_pt/colors）
    merged = {**DEFAULT_PROFILE}
    if "fonts" in user:
        merged["fonts"] = {**DEFAULT_PROFILE["fonts"], **user["fonts"]}
    if "sizes_pt" in user:
        merged["sizes_pt"] = {**DEFAULT_PROFILE["sizes_pt"], **user["sizes_pt"]}
    if "colors" in user:
        merged["colors"] = {**DEFAULT_PROFILE["colors"], **user["colors"]}
    if "footer_text" in user:
        merged["footer_text"] = user["footer_text"]
    return merged


def parse_md(text):
    """解析 markdown → list of {kind, ...}
    kind: title_main / title_sub / divider / section_h1 / clause_lead / bullet / body / signature
    """
    lines = text.splitlines()
    items = []
    main_title = sub_title = ""
    sub_title_set = False  # [Fix-1] 只取第一个 `## ` 为副标题，后续都进 section_h1
    for raw in lines:
        line = raw.rstrip()
        if not line:
            continue
        if line.startswith("# "):
            main_title = line[2:].strip()
            continue
        if line.startswith("## "):
            value = line[3:].strip()
            if not sub_title_set:
                sub_title = value
                sub_title_set = True
            else:
                items.append({"kind": "section_h1", "text": value})
            continue
        if line.startswith("### "):
            items.append({"kind": "section_h1", "text": line[4:].strip()})
            continue
        if line.startswith("**") and line.endswith("**"):
            items.append({"kind": "strong", "text": line[2:-2].strip()})
            continue
        if re.match(r"^\*\*（[一二三四五六七八九十0-9]+）", line):
            # [Fix-C] 拆 num + title + tail_body；title 仅保留 `**...**` 内的标题，剩余正文追加为 body
            m = re.match(r"^\*\*(（[^）]+）)\s*((?:[^*])*?)\*\*\s*(.*)$", line)
            if m:
                num = m.group(1)
                title = m.group(2).strip()
                tail_body = m.group(3).strip()
                items.append({"kind": "clause_lead", "num": num, "title": title})
                if tail_body:
                    items.append({"kind": "body", "text": tail_body})
                continue
        # [Fix-E] 拆 prefix + 剩余，处理 `**一是xxx。** yyy zzz` 这种格式
        if line.startswith("**一是") or line.startswith("**二是") or line.startswith("**三是") or \
           line.startswith("**四是") or line.startswith("**五是") or line.startswith("**六是") or \
           line.startswith("**七是") or line.startswith("**八是") or line.startswith("**九是"):
            m = re.match(r"^\*\*((?:[一二三四五六七八九]是)[^*]*?)\*\*\s*(.*)$", line)
            if m:
                text = (m.group(1) + " " + m.group(2)).strip()
            else:
                text = line.strip().lstrip("*").rstrip("*").strip()
            items.append({"kind": "numbered_para", "text": text})
            continue
        if line == "---":
            items.append({"kind": "divider"})
            continue
        # 正文（含 `同志们：` / `当前，` 等）
        items.append({"kind": "body", "text": line.strip()})
    return {
        "main_title": main_title,
        "sub_title": sub_title,
        "items": items,
    }


# ============================================================
# 文字 / 字体处理
# ============================================================
def strip_md_marks(text):
    """[Fix-2/7] 去掉段落里残留的 markdown 标记 ** / `"""
    if not text:
        return text
    # 去配对的 strong：**xxx** → xxx（中间无 *）
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    # 去孤立的 ** / *（开头 / 结尾 / 中间）
    text = re.sub(r"\*+", "", text)
    # 去单点 `code`
    text = re.sub(r"`([^`]+)`", r"\1", text)
    return text.strip()


def split_into_paragraphs(text):
    """[Fix-3] 把含换行的字符串拆成多段，返回 list[str]"""
    if not text:
        return [""]
    parts = [p.strip() for p in text.split("\n") if p.strip()]
    return parts if parts else [""]


# ============================================================
# 文字 / 字体处理
# ============================================================
def set_run_font(run, font_name, size_pt=None, bold=None, color=None):
    """同时设置中英文/东亚字体"""
    run.font.name = font_name
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)
    # 关键：设置东亚字体（中文）
    rPr = run._r.get_or_add_rPr()
    # 删除已存在的 eastAsia
    for ea in rPr.findall(qn("a:ea")):
        rPr.remove(ea)
    # 添加新的 eastAsia
    from lxml import etree
    ea = etree.SubElement(rPr, qn("a:ea"))
    ea.set("typeface", font_name)


def add_textbox(slide, x, y, w, h, text="", font="body", size=18,
                color="1A1A1A", bold=False, align="left", anchor="top",
                line_spacing=1.3, profile=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2); tf.margin_right = Pt(2)
    tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
    anchor_map = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}
    tf.vertical_anchor = anchor_map.get(anchor, MSO_ANCHOR.TOP)
    align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    p = tf.paragraphs[0]
    p.alignment = align_map.get(align, PP_ALIGN.LEFT)
    p.line_spacing = line_spacing
    if profile:
        font_name = profile["fonts"].get(font, font)
    else:
        font_name = font
    run = p.add_run()
    run.text = text
    set_run_font(run, font_name, size, bold, color)
    return tb


def add_filled_rect(slide, x, y, w, h, fill="1F5FA8", line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = RGBColor.from_string(line)
    shape.shadow.inherit = False
    return shape


# ============================================================
# 通用页面元素
# ============================================================
def add_header_band(slide, profile, label=""):
    """顶部蓝色色条 + 集团标识"""
    w = profile["page_w_in"]
    add_filled_rect(slide, 0, 0, w, 0.45, fill=profile["colors"]["primary"])
    # 装饰小红条
    add_filled_rect(slide, 0, 0.45, w, 0.06, fill=profile["colors"]["accent"])
    if label:
        add_textbox(slide, 0.5, 0.08, w - 1.0, 0.32,
                    text=label, font="number", size=11,
                    color=profile["colors"]["white"], bold=True,
                    align="left", anchor="middle", profile=profile)


def add_footer(slide, profile, page_no, total):
    """底部页脚：集团名 + 页码"""
    w = profile["page_w_in"]
    h = profile["page_h_in"]
    # 灰线
    add_filled_rect(slide, 0.5, h - 0.45, w - 1.0, 0.015, fill=profile["colors"]["muted"])
    add_textbox(slide, 0.5, h - 0.38, w - 2.5, 0.3,
                text=profile["footer_text"], font="body",
                size=profile["sizes_pt"]["footer"],
                color=profile["colors"]["muted"], align="left",
                anchor="middle", profile=profile)
    add_textbox(slide, w - 2.0, h - 0.38, 1.5, 0.3,
                text=f"—— {page_no} / {total} ——", font="body",
                size=profile["sizes_pt"]["footer"],
                color=profile["colors"]["muted"], align="right",
                anchor="middle", profile=profile)


# ============================================================
# 各页面构造
# ============================================================
def make_cover(prs, parsed, profile, page_no, total):
    """封面页"""
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    w, h = profile["page_w_in"], profile["page_h_in"]

    # 顶部蓝色色块
    add_filled_rect(slide, 0, 0, w, 2.6, fill=profile["colors"]["primary"])
    # 装饰斜条
    add_filled_rect(slide, 0, 2.6, w, 0.08, fill=profile["colors"]["accent"])
    # 底部色块（淡）
    add_filled_rect(slide, 0, h - 0.9, w, 0.9, fill=profile["colors"]["light"])

    # 集团 LOGO 占位（用文字代替）
    add_textbox(slide, 0.5, 0.5, 6.0, 0.45,
                text="SZ ENV WATER · 深圳环境水务集团", font="number",
                size=12, color=profile["colors"]["white"],
                bold=True, align="left", anchor="middle", profile=profile)
    # 右上：会议标签
    add_textbox(slide, w - 5.5, 0.5, 5.0, 0.45,
                text="WORK REPORT 2026 H1", font="number",
                size=11, color=profile["colors"]["white"],
                bold=True, align="right", anchor="middle", profile=profile)

    # 主标题
    add_textbox(slide, 0.8, 1.0, w - 1.6, 1.0,
                text=parsed["main_title"], font="title_main",
                size=profile["sizes_pt"]["title_main"],
                color=profile["colors"]["white"], bold=True,
                align="center", anchor="middle", profile=profile)

    # 副标题
    add_textbox(slide, 0.8, 2.0, w - 1.6, 0.5,
                text=parsed["sub_title"], font="title_sub",
                size=profile["sizes_pt"]["title_sub"],
                color=profile["colors"]["white"], bold=True,
                align="center", anchor="middle", profile=profile)

    # 落款行
    add_textbox(slide, 0.8, 3.4, w - 1.6, 0.5,
                text="集团党委书记、董事长  龚利民", font="body",
                size=20, color=profile["colors"]["dark"],
                bold=False, align="center", anchor="middle", profile=profile)

    add_textbox(slide, 0.8, 3.95, w - 1.6, 0.5,
                text="（2026年9月10日）", font="body",
                size=18, color=profile["colors"]["muted"],
                align="center", anchor="middle", profile=profile)

    # 底部信息条
    add_textbox(slide, 0.5, h - 0.75, w - 1.0, 0.4,
                text="中共深圳环境水务集团有限公司委员会", font="body",
                size=14, color=profile["colors"]["primary"],
                bold=True, align="center", anchor="middle", profile=profile)

    # 页脚
    add_footer(slide, profile, page_no, total)


def make_toc(prs, profile, page_no, total):
    """目录页"""
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    w, h = profile["page_w_in"], profile["page_h_in"]

    add_header_band(slide, profile, "目  录  CONTENTS")

    # 顶部大标题
    add_textbox(slide, 0.5, 0.9, w - 1.0, 0.8,
                text="本次汇报主要内容", font="title_main",
                size=36, color=profile["colors"]["primary"],
                bold=True, align="center", anchor="middle", profile=profile)

    # 列表
    items = [
        ("01", "上半年工作回顾", "5条成效  ·  3条不足"),
        ("02", "下半年工作要求", "7条要求  ·  重点突破"),
        ("03", "结语", "团结拼搏  ·  勇毅前行"),
    ]
    y0 = 2.2
    for i, (num, t, sub) in enumerate(items):
        yy = y0 + i * 1.4
        # 编号块
        add_filled_rect(slide, 0.8, yy, 1.0, 1.0, fill=profile["colors"]["primary"])
        add_textbox(slide, 0.8, yy, 1.0, 1.0,
                    text=num, font="number", size=32,
                    color=profile["colors"]["white"], bold=True,
                    align="center", anchor="middle", profile=profile)
        # 文字
        add_textbox(slide, 2.0, yy, w - 2.8, 0.55,
                    text=t, font="title_sub", size=22,
                    color=profile["colors"]["dark"], bold=True,
                    align="left", anchor="bottom", profile=profile)
        add_textbox(slide, 2.0, yy + 0.55, w - 2.8, 0.45,
                    text=sub, font="body", size=16,
                    color=profile["colors"]["muted"],
                    align="left", anchor="top", profile=profile)

    add_footer(slide, profile, page_no, total)


def make_section_divider(prs, profile, page_no, total, label, part="01"):
    """章节封面（一级章节）"""
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    w, h = profile["page_w_in"], profile["page_h_in"]

    # 整版蓝色
    add_filled_rect(slide, 0, 0, w, h, fill=profile["colors"]["primary"])
    # 装饰小红条
    add_filled_rect(slide, 0, h - 0.4, w, 0.06, fill=profile["colors"]["accent"])

    # 大标签
    add_textbox(slide, 1.0, 2.0, w - 2.0, 1.0,
                text=label, font="title_main",
                size=44, color=profile["colors"]["white"], bold=True,
                align="center", anchor="middle", profile=profile)

    # 下方细字
    add_textbox(slide, 1.0, 3.4, w - 2.0, 0.5,
                text="—— 集团2026年上半年工作总结 ——", font="body",
                size=18, color=profile["colors"]["white"],
                align="center", anchor="middle", profile=profile)

    # 顶部副标
    add_textbox(slide, 0.5, 0.5, w - 1.0, 0.4,
                text=f"PART {part}", font="number",
                size=14, color=profile["colors"]["accent"], bold=True,
                align="right", anchor="middle", profile=profile)

    # 页脚
    add_footer(slide, profile, page_no, total)


def add_paragraph_text(tf, text, font_name, size_pt, color, bold=False, align=PP_ALIGN.LEFT, line_spacing=1.5):
    """[Fix-3] 把 text（含 \n）写入 text_frame，每行一个 para（替代把整段塞进单 run）"""
    if not text:
        return
    paragraphs = [p.strip() for p in text.split("\n") if p.strip()]
    if not paragraphs:
        paragraphs = [""]
    for i, para_text in enumerate(paragraphs):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        run = p.add_run()
        run.text = para_text
        set_run_font(run, font_name, size_pt, bold, color)


def make_achievement_page(prs, profile, page_no, total, num, title, body):
    """成效页：大编号 + 标题 + 4-6 个 bullet"""
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    w, h = profile["page_w_in"], profile["page_h_in"]

    add_header_band(slide, profile, "上半年工作成效")

    # 左侧大编号块
    add_filled_rect(slide, 0.5, 1.0, 3.0, 5.0, fill=profile["colors"]["primary"])
    add_filled_rect(slide, 0.5, 1.0, 0.18, 5.0, fill=profile["colors"]["accent"])
    # 巨型数字
    big = num.replace("是", "") if num.startswith("一") else num
    add_textbox(slide, 0.5, 1.4, 3.0, 1.8,
                text=big, font="number", size=120,
                color=profile["colors"]["white"], bold=True,
                align="center", anchor="middle", profile=profile)
    add_textbox(slide, 0.5, 3.4, 3.0, 0.6,
                text="成效·ACHIEVEMENT", font="number", size=14,
                color=profile["colors"]["accent"], bold=True,
                align="center", anchor="top", profile=profile)
    add_textbox(slide, 0.5, 4.0, 3.0, 0.6,
                text=f"PAGE {page_no:02d} / {total:02d}", font="number", size=11,
                color=profile["colors"]["light"],
                align="center", anchor="top", profile=profile)

    # 右侧标题
    add_textbox(slide, 4.0, 1.0, w - 4.5, 0.7,
                text=title, font="title_sub", size=28,
                color=profile["colors"]["primary"], bold=True,
                align="left", anchor="middle", profile=profile)
    add_filled_rect(slide, 4.0, 1.75, 1.5, 0.05, fill=profile["colors"]["accent"])

    # 右侧正文（去掉前缀"一是"，strip markdown，每段单独成行）
    body_clean = strip_md_marks(body)
    m2 = re.match(r"^[一二三四五六七八九]是(.*)", body_clean)
    if m2:
        body_clean = m2.group(1)
    # [Fix-3] 按段落（\n 已由 split_into_paragraphs 处理）+ 句末标点切分
    raw_paras = split_into_paragraphs(body_clean)
    bullet_lines = []
    for raw in raw_paras:
        # 每段内按"。"或"；"切
        parts = re.split(r"(?<=[。；])", raw)
        parts = [p.strip() for p in parts if p.strip()]
        bullet_lines.extend(parts)
    # [Fix-3] 限制条数 + 转 bullet 串
    bullet_lines = bullet_lines[:6]
    bullet_text = "\n".join([f"▸ {p}" for p in bullet_lines]) if bullet_lines else body_clean

    tb = slide.shapes.add_textbox(Inches(4.0), Inches(1.95),
                                   Inches(w - 4.5), Inches(h - 2.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2); tf.margin_right = Pt(2)
    tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    add_paragraph_text(tf, bullet_text, profile["fonts"]["body"], 17,
                       profile["colors"]["text"], bold=False,
                       align=PP_ALIGN.LEFT, line_spacing=1.5)

    add_footer(slide, profile, page_no, total)


def make_clause_page(prs, profile, page_no, total, num, title, body):
    """条款页：左编号块 + 右上大字 + 右下要点"""
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    w, h = profile["page_w_in"], profile["page_h_in"]

    add_header_band(slide, profile, "下半年工作要求")

    # [Fix-6] 左侧编号块：根据 num 字符数自适应字号（5 字符 36pt 起步）
    num_len = len(num)
    if num_len <= 3:
        num_size = 54
    elif num_len == 4:
        num_size = 46
    else:  # 5 字符如"（十）"或长编号
        num_size = 36
    add_filled_rect(slide, 0.5, 1.0, 3.0, 5.0, fill=profile["colors"]["accent"])
    add_filled_rect(slide, 0.5, 1.0, 0.18, 5.0, fill=profile["colors"]["primary"])
    add_textbox(slide, 0.5, 1.4, 3.0, 1.6,
                text=num, font="title_sub", size=num_size,
                color=profile["colors"]["white"], bold=True,
                align="center", anchor="middle", profile=profile)
    add_textbox(slide, 0.5, 3.1, 3.0, 0.6,
                text="要求·REQUIREMENT", font="number", size=14,
                color=profile["colors"]["white"], bold=True,
                align="center", anchor="top", profile=profile)
    add_textbox(slide, 0.5, 4.0, 3.0, 0.6,
                text=f"PAGE {page_no:02d} / {total:02d}", font="number", size=11,
                color=profile["colors"]["white"],
                align="center", anchor="top", profile=profile)

    # 右侧标题（楷体，仿"条款"风格）
    add_textbox(slide, 4.0, 1.0, w - 4.5, 0.9,
                text=title, font="clause", size=26,
                color=profile["colors"]["primary"], bold=True,
                align="left", anchor="middle", profile=profile)
    add_filled_rect(slide, 4.0, 1.95, 1.5, 0.05, fill=profile["colors"]["accent"])

    # 右侧正文：[Fix-2] strip markdown 残留 + [Fix-3] 多段化
    body_clean = strip_md_marks(body)
    # 移除开头的编号部分（如"（一）"）
    m2 = re.match(r"^（[^）]+）\s*(.*)", body_clean)
    if m2:
        body_clean = m2.group(1)
    raw_paras = split_into_paragraphs(body_clean)
    bullet_lines = []
    for raw in raw_paras:
        parts = re.split(r"(?<=[。；])", raw)
        parts = [p.strip() for p in parts if p.strip()]
        bullet_lines.extend(parts)
    bullet_lines = bullet_lines[:6]
    bullet_text = "\n".join([f"▸ {p}" for p in bullet_lines]) if bullet_lines else body_clean

    tb = slide.shapes.add_textbox(Inches(4.0), Inches(2.15),
                                   Inches(w - 4.5), Inches(h - 2.7))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2); tf.margin_right = Pt(2)
    tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    add_paragraph_text(tf, bullet_text, profile["fonts"]["body"], 17,
                       profile["colors"]["text"], bold=False,
                       align=PP_ALIGN.LEFT, line_spacing=1.5)

    add_footer(slide, profile, page_no, total)


def make_content_page(prs, profile, page_no, total, num_label, title, body):
    """正文页（左编号 + 右要点）"""
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    w, h = profile["page_w_in"], profile["page_h_in"]

    add_header_band(slide, profile, "上半年工作回顾" if "是" in num_label or "工作" in title else "下半年工作要求")

    # 左侧编号块
    add_filled_rect(slide, 0.5, 1.0, 2.6, 4.5, fill=profile["colors"]["light"])
    add_filled_rect(slide, 0.5, 1.0, 0.15, 4.5, fill=profile["colors"]["accent"])
    add_textbox(slide, 0.5, 1.2, 2.6, 1.0,
                text=num_label.split("（")[0] if "（" in num_label else num_label[:2],
                font="number", size=44, color=profile["colors"]["primary"],
                bold=True, align="center", anchor="middle", profile=profile)
    add_textbox(slide, 0.5, 2.2, 2.6, 0.5,
                text=num_label[-1] if "（" in num_label else "",
                font="number", size=20, color=profile["colors"]["accent"],
                bold=True, align="center", anchor="top", profile=profile)
    add_textbox(slide, 0.5, 2.7, 2.6, 0.6,
                text="PAGE", font="number", size=11,
                color=profile["colors"]["muted"],
                align="center", anchor="top", profile=profile)

    # 右侧标题 + 正文
    add_textbox(slide, 3.5, 1.0, w - 4.0, 0.8,
                text=title, font="title_sub", size=28,
                color=profile["colors"]["primary"], bold=True,
                align="left", anchor="middle", profile=profile)
    # 装饰线
    add_filled_rect(slide, 3.5, 1.85, 1.5, 0.04, fill=profile["colors"]["accent"])

    # [Fix-2/3] 正文去 markdown + 多段化
    body_clean = strip_md_marks(body)
    tb = slide.shapes.add_textbox(Inches(3.5), Inches(2.05),
                                   Inches(w - 4.0), Inches(4.0))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2); tf.margin_right = Pt(2)
    tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    add_paragraph_text(tf, body_clean, profile["fonts"]["body"], 18,
                       profile["colors"]["text"], bold=False,
                       align=PP_ALIGN.LEFT, line_spacing=1.5)

    add_footer(slide, profile, page_no, total)


def make_simple_text_page(prs, profile, page_no, total, title, body, label="全文要点"):
    """纯文字页（开场致辞、结语等）"""
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    w, h = profile["page_w_in"], profile["page_h_in"]

    add_header_band(slide, profile, label)

    # 顶部大标题
    add_textbox(slide, 0.5, 0.9, w - 1.0, 0.7,
                text=title, font="title_sub", size=28,
                color=profile["colors"]["primary"], bold=True,
                align="center", anchor="middle", profile=profile)
    add_filled_rect(slide, w / 2 - 1.0, 1.7, 2.0, 0.05, fill=profile["colors"]["accent"])

    # [Fix-2/3] 正文去 markdown + 多段化
    body_clean = strip_md_marks(body)
    tb = slide.shapes.add_textbox(Inches(1.0), Inches(2.1),
                                   Inches(w - 2.0), Inches(h - 3.0))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(2); tf.margin_right = Pt(2)
    tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    add_paragraph_text(tf, body_clean, profile["fonts"]["body"], 18,
                       profile["colors"]["text"], bold=False,
                       align=PP_ALIGN.LEFT, line_spacing=1.6)

    add_footer(slide, profile, page_no, total)


def make_insufficient_page(prs, profile, page_no, total, items):
    """不足页：3 列并列展示"""
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    w, h = profile["page_w_in"], profile["page_h_in"]

    add_header_band(slide, profile, "工作中的不足")

    add_textbox(slide, 0.5, 0.9, w - 1.0, 0.7,
                text="工作中的不足", font="title_sub", size=30,
                color=profile["colors"]["primary"], bold=True,
                align="center", anchor="middle", profile=profile)
    add_filled_rect(slide, w / 2 - 1.0, 1.7, 2.0, 0.05, fill=profile["colors"]["accent"])

    # 3 列卡片
    col_w = (w - 1.5) / 3
    y0 = 2.0
    h_card = 4.2
    for i, it in enumerate(items):
        num, title = extract_numbered_title(it)
        x0 = 0.5 + i * (col_w + 0.15)

        # 卡片底
        add_filled_rect(slide, x0, y0, col_w, h_card, fill=profile["colors"]["light"])
        # 顶部色条
        add_filled_rect(slide, x0, y0, col_w, 0.12, fill=profile["colors"]["accent"])
        # 编号
        add_textbox(slide, x0, y0 + 0.2, col_w, 0.7,
                text=num.split("（")[0] if "（" in num else num[:2],
                font="number", size=42, color=profile["colors"]["primary"],
                bold=True, align="center", anchor="middle", profile=profile)
        # 标题
        add_textbox(slide, x0 + 0.15, y0 + 1.0, col_w - 0.3, 1.0,
                text=title, font="title_sub", size=18,
                color=profile["colors"]["dark"], bold=True,
                align="center", anchor="top", line_spacing=1.2,
                profile=profile)
        # 正文：[Fix-2] 去掉前缀 + strip markdown
        body = it["text"]
        m = re.match(r"^[一二三四五六七八九]是(.*)", body)
        body_text = m.group(1) if m else body
        body_text = body_text.rstrip("。. ").strip()
        body_text = strip_md_marks(body_text)
        add_textbox(slide, x0 + 0.2, y0 + 2.1, col_w - 0.4, 2.0,
                text=body_text, font="body", size=14,
                color=profile["colors"]["text"],
                align="left", anchor="top", line_spacing=1.4,
                profile=profile)

    add_footer(slide, profile, page_no, total)


def make_signature_page(prs, profile, page_no, total):
    """落款页"""
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    w, h = profile["page_w_in"], profile["page_h_in"]

    # 整版浅蓝底
    add_filled_rect(slide, 0, 0, w, h, fill=profile["colors"]["light"])

    # 顶部色条
    add_filled_rect(slide, 0, 0, w, 0.45, fill=profile["colors"]["primary"])
    add_filled_rect(slide, 0, 0.45, w, 0.06, fill=profile["colors"]["accent"])

    # 中间段落
    add_textbox(slide, 1.0, 2.5, w - 2.0, 1.0,
                text="感谢聆听", font="title_main", size=54,
                color=profile["colors"]["primary"], bold=True,
                align="center", anchor="middle", profile=profile)

    add_textbox(slide, 1.0, 3.7, w - 2.0, 0.5,
                text="深圳环境水务集团 · 2026年上半年工作总结推进会议",
                font="body", size=18, color=profile["colors"]["dark"],
                align="center", anchor="middle", profile=profile)

    add_textbox(slide, 1.0, 4.4, w - 2.0, 0.4,
                text="集团党委书记、董事长  龚利民   |   2026年9月10日",
                font="body", size=16, color=profile["colors"]["muted"],
                align="center", anchor="middle", profile=profile)

    # 底部装饰
    add_filled_rect(slide, w / 2 - 1.5, 5.2, 3.0, 0.06, fill=profile["colors"]["accent"])

    add_footer(slide, profile, page_no, total)


# ============================================================
# 主流程
# ============================================================
def extract_numbered_title(item):
    """从 numbered_para / clause_lead / strong 中提取 (num, title)"""
    if item["kind"] == "numbered_para":
        # "一是经营指标稳中有进。"
        text = item["text"].rstrip("。. ")
        for prefix in ["一是","二是","三是","四是","五是","六是","七是","八是","九是"]:
            if text.startswith(prefix):
                return prefix, text[len(prefix):]
        return "·", text
    if item["kind"] == "clause_lead":
        # "（一）坚定不移完成全年固投81亿元目标。"
        num = item["num"]
        title = item["title"].rstrip("。. ")
        return num, title
    if item["kind"] == "strong":
        return "·", item["text"].rstrip("。. ")
    return "·", item.get("text", "")


def _find_opening_closing(items):
    """[Fix-5] 智能定位开场致辞与结语正文。

    启发式：
    - 开场致辞：第 1 个 divider 之后、第 1 个 section_h1 之前的第一个长 body（≥ 30 字）
    - 结语正文：最后 1 个 section_h1 之后的第一个 body（剥离 leading `**...**` 标记）

    兜底：取 items 中第一个 / 最后一个长 body
    """
    divs = [i for i, x in enumerate(items) if x["kind"] == "divider"]
    secs = [i for i, x in enumerate(items) if x["kind"] == "section_h1"]

    opening = ""
    if divs and secs:
        lo, hi = divs[0], secs[0]
        candidates = [items[i]["text"] for i in range(lo, hi)
                      if items[i]["kind"] == "body" and len(items[i]["text"]) >= 30]
        if candidates:
            opening = max(candidates, key=len)

    closing = ""
    if secs:
        last_sec = secs[-1]
        # 找最后一个 section_h1 之后第一个 divider 之前的所有 body（剥离 leading strong 标记）
        end_pos = len(items)
        for k in range(last_sec + 1, len(items)):
            if items[k]["kind"] == "divider":
                end_pos = k
                break
        candidates = [items[i]["text"] for i in range(last_sec + 1, end_pos)
                      if items[i]["kind"] == "body"]
        # 优先含结语关键词
        for t in candidates:
            if any(kw in t for kw in ("同志们", "奋进", "谱写", "不进则退", "锐始者", "功成不必在我", "新篇章")):
                closing = strip_md_marks(t)
                break
        if not closing and candidates:
            closing = strip_md_marks(max(candidates, key=len))

    # 兜底：取 items 中第一个 / 最后一个长 body
    if not opening:
        long_bodies = [items[i]["text"] for i, x in enumerate(items)
                       if x["kind"] == "body" and len(x["text"]) >= 30]
        if long_bodies:
            opening = long_bodies[0]
    if not closing:
        long_bodies = [items[i]["text"] for i, x in enumerate(items)
                       if x["kind"] == "body" and len(x["text"]) >= 30]
        if long_bodies:
            closing = strip_md_marks(long_bodies[-1])

    return opening, closing


def build_pptx(input_md, output_pptx, profile):
    text = Path(input_md).read_text(encoding="utf-8")
    parsed = parse_md(text)

    prs = Presentation()
    prs.slide_width = Inches(profile["page_w_in"])
    prs.slide_height = Inches(profile["page_h_in"])

    items = parsed["items"]
    numbered = [x for x in items if x["kind"] == "numbered_para"]
    clauses = [x for x in items if x["kind"] == "clause_lead"]

    # [Fix-5] 用语义定位替换 body_idx[0] / body_idx[-1]
    opening, closing = _find_opening_closing(items)

    # 找第三个及以后的 "一是..三是" / "一是..九是" 中间算不足部分
    # 由于本次解析按 markdown 顺序——成效是 5 个 numbered_para，
    # 然后可能还有 body 段（在"工作中不足"标题下的 strong / numbered_para）。
    # 简化策略：numbered_para 中第 6、7、8 个算不足
    insufficient = numbered[5:8] if len(numbered) >= 8 else numbered[5:]
    achievements = numbered[:5] if len(numbered) >= 5 else numbered

    plan = [
        ("cover", {}),
        ("toc", {}),
        ("section", {"label": "一、上半年工作回顾", "part": "01"}),
        ("text", {"title": "开场致辞", "body": opening, "label": "上半年工作回顾"}),
    ]
    # 5 条成效
    for it in achievements:
        num, title = extract_numbered_title(it)
        plan.append(("achievement", {"num": num, "title": title, "body": it["text"]}))
    # 不足合并一页（3 列）
    if insufficient:
        plan.append(("insufficient", {"items": insufficient}))
    # 第二章封面
    plan.append(("section", {"label": "二、下半年工作要求", "part": "02"}))
    # 7 条要求（clause_lead 仅有 num + title，需从前后文找正文）
    # [Fix-4] 用 items_full 中 clause 的真实索引，而不是 enumerate(clauses) 的 idx
    items_full = parsed["items"]
    clause_positions = [i for i, x in enumerate(items_full) if x["kind"] == "clause_lead"]
    for real_idx in clause_positions:
        it = items_full[real_idx]
        num = it["num"]
        title = it["title"]
        # 收集下一个 clause / numbered_para / section_h1 / divider 之前的所有 body / strong
        # [Fix-D] break 条件加 section_h1 + divider，避免越界吞掉后续章节内容
        body_parts = []
        for j in range(real_idx + 1, len(items_full)):
            nxt = items_full[j]
            if nxt["kind"] in ("clause_lead", "numbered_para", "section_h1", "divider"):
                break
            if nxt["kind"] == "body":
                body_parts.append(nxt["text"])
            elif nxt["kind"] == "strong":
                body_parts.append(nxt["text"])
        body = "\n".join(body_parts)
        plan.append(("clause", {"num": num, "title": title, "body": body}))
    # 结语
    plan.append(("text", {"title": "结  语", "body": closing, "label": "下半年工作要求"}))
    plan.append(("signature", {}))

    total = len(plan)

    for idx, (kind, kw) in enumerate(plan, start=1):
        if kind == "cover":
            make_cover(prs, parsed, profile, idx, total)
        elif kind == "toc":
            make_toc(prs, profile, idx, total)
        elif kind == "section":
            make_section_divider(prs, profile, idx, total, kw["label"], kw.get("part", "01"))
        elif kind == "achievement":
            make_achievement_page(prs, profile, idx, total, kw["num"], kw["title"], kw["body"])
        elif kind == "insufficient":
            make_insufficient_page(prs, profile, idx, total, kw["items"])
        elif kind == "clause":
            make_clause_page(prs, profile, idx, total, kw["num"], kw["title"], kw["body"])
        elif kind == "text":
            make_simple_text_page(prs, profile, idx, total,
                                  kw["title"], kw.get("body", ""), kw.get("label", ""))
        elif kind == "signature":
            make_signature_page(prs, profile, idx, total)

    Path(output_pptx).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_pptx)
    return total


def main():
    ap = argparse.ArgumentParser(description="Markdown → PPTX (领导风格)")
    ap.add_argument("--input", required=True, help="输入 markdown 源稿")
    ap.add_argument("--output", required=True, help="输出 pptx")
    ap.add_argument("--leader", default="龚利民", help="领导姓名（档案标识）")
    ap.add_argument("--profile", default=None, help="profile yaml 路径（可选）")
    ap.add_argument("--ratio", default="16:9", choices=["16:9", "4:3"])
    args = ap.parse_args()

    profile = load_profile(args.profile)
    if args.ratio == "4:3":
        profile["page_w_in"] = 10.0
        profile["page_h_in"] = 7.5

    total = build_pptx(args.input, args.output, profile)
    print(f"✅ PPTX 生成完毕：{args.output}  共 {total} 页")


if __name__ == "__main__":
    main()