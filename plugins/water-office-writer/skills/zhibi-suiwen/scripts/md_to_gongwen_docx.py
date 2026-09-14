#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md_to_gongwen_docx.py — 把 Markdown 报告转成集团公文档式 docx

用法:
    python md_to_gongwen_docx.py --input report.md --output report.docx --leader 龚利民
    python md_to_gongwen_docx.py --input report.md --output out.docx --leader 方总 \
        --clause-scope lead --normalize-punct --strict-leader
    python md_to_gongwen_docx.py --input speech.md --output out.docx --leader 龚利民 --doc-type speech

版式来源（优先级 高 → 低）:
    ① --profile 显式指定的风格档案 YAML
    ② skill 内置 assets/profile_<领导>_<文种>_v*.yaml
       （文种=总结/致辞/呈报，按 --doc-type 对应匹配）
       → 其次兼容旧命名 assets/profile_<领导>_v*.yaml（取最高版本）
    ③ 本脚本内置的实测兜底表 FALLBACK_PROFILES
    ④ 未识别领导 → 默认版式（可通过 --strict-leader 改为报错）

    龚利民（董事长）  : 正文 仿宋_GB2312 小二 18pt / 条款 楷体 18pt / 缩进 1.28cm / 边距 37,35,26,26
    方琳（总裁）  : 正文 仿宋_GB2312 三号 16pt / 条款 楷体 16pt / 缩进 1.13cm / 边距 37,35,28,26
    共同: 主标题 方正小标宋简体 22pt 居中 / 副标题 黑体 18pt 居中 / 一级标题 黑体 / A4 / 行距固定 28pt / 页码页脚居中

文种（--doc-type，默认 report）:
    report 总结/汇报类：四字对仗标题 + 黑体副标题 + 文末右对齐落款 + 页码
    speech 致辞/讲话类（题注式，实测自 GIIC鸿蒙大会/广东AI大会等 5 份样本）：
        · 主标题「在XXX上的致辞」句式（# ）
        · ## 署名行（职务+姓名）   -> 楷体 16pt 居中（题注式，替代文末落款）
        · ## （YYYY年M月D日）      -> 仿宋 16pt 居中 日期行
        · 正文无黑体标题层级，分点内嵌「一是/二是/三是」
        · **（共N字，约M分钟）**   -> 楷体 16pt 右对齐 字数注记（缺失时自动补算）
        · 无页码（页脚空）、文末不落款
    baogao 呈报类（对外汇报/专题汇报/经验材料，实测自水利部报告、鸿蒙生态报告等 4 份样本）：
        · 主标题「（单位）关于XX（情况）的报告」句式（# ），小标宋 22pt 居中
        · 主送机关「XX部：」/称呼「尊敬的XX市长：」-> 顶格不缩进（普通段落即可识别）
        · 正文仿宋_GB2312 16pt（不是总结的 18pt！）/ 缩进 1.13cm
        · 「一、二、三、」独立短行   -> 黑体 16pt（呈报模式自动识别）
        · （一）（二）条款          -> 楷体 16pt（同总结）
        · 结尾「专此报告。」+（可选）「附件：XXX」普通段落
        · **单位全称** + **日期**    -> 仿宋 16pt 右对齐（单位落款，非个人姓名）
        · 有页码（页脚居中 PAGE 域）

Markdown 约定:
    # 主标题              -> 居中 方正小标宋简体 22pt
    ## 副标题 / 题注       -> 居中 黑体 18pt
    ### 一级标题          -> 黑体，与正文同级字号，2 字符缩进
    #### 四级标题         -> 楷体加粗，与正文同级字号（新增）
    （一）/(一) 开头       -> 楷体条款；--clause-scope lead 时仅引语楷体、余下仿宋（新增）
    1. 开头               -> 三级层次，仿宋 + 2 字符缩进；--h3-bold 时加粗（新增）
    - / * 开头            -> 无序列表项，剥离标记转仿宋正文（新增）
    XX：结尾（≤60字）      -> 称呼语，顶格不缩进，带话题白名单防误判（已修正）
    **整行**              -> 落款，右对齐
    | a | b | 连续多行     -> 真表格，三线表样式（新增）
    **粗体** / *斜体* / `码` -> 行内富文本，去除标记（新增）
    > ...                -> 引用块，仅 markdown 保留，不写入 docx
    ---                  -> 空行

依赖:
    pip install python-docx pyyaml
"""
import argparse
import glob
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT_TITLE = "方正小标宋简体"
FONT_BODY = "仿宋_GB2312"
FONT_KAI = "楷体"
FONT_HEI = "黑体"
FONT_LATIN = "Times New Roman"

# 兜底版式（实测值，单位 mm / pt / cm）——档案解析失败时使用
FALLBACK_PROFILES = {
    "龚利民": dict(body_pt=18, indent_cm=1.28, margin=(37, 35, 26, 26),
                   h1_font=FONT_HEI, h1_pt=18, clause_font=FONT_KAI,
                   body_font=FONT_BODY, title_pt=22, subtitle_pt=18,
                   line_pt=28, page_number="bottom-center"),
    "方琳":   dict(body_pt=16, indent_cm=1.13, margin=(37, 35, 28, 26),
                   h1_font=FONT_HEI, h1_pt=16, clause_font=FONT_KAI,
                   body_font=FONT_BODY, title_pt=22, subtitle_pt=18,
                   line_pt=28, page_number="bottom-center"),
}
ALIASES = {
    "龚董": "龚利民", "董事长": "龚利民", "gong": "龚利民", "gonglimin": "龚利民",
    "方总": "方琳", "总裁": "方琳", "fang": "方琳", "fanglin": "方琳",
}
DEFAULT_PROFILE = dict(body_pt=16, indent_cm=1.13, margin=(37, 35, 28, 26),
                       h1_font=FONT_HEI, h1_pt=16, clause_font=FONT_KAI,
                       body_font=FONT_BODY, title_pt=22, subtitle_pt=18,
                       line_pt=28, page_number="bottom-center")

# ---- 文种维度（--doc-type）：总结 report / 致辞讲话 speech / 呈报 baogao ----
DOC_TYPE_ALIASES = {
    "report": "report", "总结": "report", "报告": "report", "汇报": "report",
    "speech": "speech", "致辞": "speech", "讲话": "speech", "发言": "speech",
    "baogao": "baogao", "呈报": "baogao", "对外汇报": "baogao", "专题汇报": "baogao",
    "呈报材料": "baogao", "经验材料": "baogao", "经验交流": "baogao", "专报": "baogao",
}
# 致辞题注区识别：日期行（（YYYY年M月D日））与文末字数注记（（共N字，约M分钟））
DATE_LINE_RE = re.compile(r"^（\d{4}年\d{1,2}月\d{1,2}日）$")
WORDCOUNT_NOTE_RE = re.compile(r"^（共.{1,10}字[，,].{1,10}分钟）$")
# speech 模式兜底：署名/日期/字数注记（无致辞档案时的默认值）
SPEECH_DEFAULTS = dict(byline_font=FONT_KAI, byline_pt=16,
                       date_font=FONT_BODY, date_pt=16,
                       sig_font=FONT_KAI, sig_pt=16)
# baogao 模式：一级标题「一、二、三、」（黑体，独立短行）与主送机关「XX部：」
LEVEL1_RE = re.compile(r"^[一二三四五六七八九十]+、")
MASTER_ORGAN_RE = re.compile(r"^[\u4e00-\u9fff]{2,14}(?:部|委|厅|局|府|政府|办公室|国资委|管理局|监管局)：$")

CLAUSE_RE = re.compile(r"^[（(][一二三四五六七八九十]+[)）]")
CLAUSE_LEAD_INLINE_RE = re.compile(r"^[（(][一二三四五六七八九十]+[)）][^。]{0,40}。")
LIST_UL_RE = re.compile(r"^\s*[-*+]\s+(.*)$")
LIST_OL_RE = re.compile(r"^\s*(\d{1,2})[.、)]\s*(.+)$")
TABLE_SEP_RE = re.compile(r"^\s*\|?[\s:|-]+\|[\s:|-]*$")

# ---- 称呼语识别：话题白名单 + 阻断词（修正原「≤25 字且以冒号结尾」的双向失准）----
SALUTATION_HINT = re.compile(
    r"(同志们|各位|尊敬的|女士们|先生们|同仁|嘉宾|领导|同事|委员|董事长|副董事长"
    r"|总裁|副总裁|总经理|局长|副局长|主任|院长|校长|书记|主席|部长|司长|处长|朋友|来宾)")
SALUTATION_BLOCK = re.compile(
    r"(如下|以下|详见|附表|附图|如下图|如下表|情况|安排|要求|标准|明细|清单|目标|指标|比例)")
NAME_LIST_RE = re.compile(r"[\u4e00-\u9fff]{2,4}([、，,][\u4e00-\u9fff]{2,4})+")

# OOXML 子元素顺序（用于安全插入，避免 Word 报文档损坏）
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
               "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
               "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
               "tblCaption", "tblDescription", "tblPrChange"]
TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
              "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
              "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"]


def _insert_ordered(parent, el, order):
    """按 OOXML schema 顺序把 el 插入 parent"""
    tag = el.tag.split("}")[-1]
    try:
        idx = order.index(tag)
    except ValueError:
        parent.append(el)
        return el
    for child in list(parent):
        ctag = child.tag.split("}")[-1]
        if ctag in order and order.index(ctag) > idx:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


# ============================================================
# 1. 版式档案解析
# ============================================================
def resolve_leader(leader: str) -> str:
    k = (leader or "").strip()
    return ALIASES.get(k, k)


def _dig(d, *keys, default=None):
    cur = d
    for k in keys:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur if cur is not None else default


def profile_from_yaml(yaml_path: Path) -> dict:
    """把风格档案 YAML 映射为落版参数字典（兼容 profile_*_v1.yaml 的嵌套结构）"""
    import yaml
    raw = yaml.safe_load(yaml_path.read_text(encoding="utf-8")) or {}
    m = _dig(raw, "format", "page", "margins_mm", default={}) or {}
    prof = dict(FALLBACK_PROFILES.get(_dig(raw, "meta", "leader_name", default=""),
                                      DEFAULT_PROFILE))
    font_map = {
        "body_font":      _dig(raw, "format", "fonts", "body", "family"),
        "body_pt":        _dig(raw, "format", "fonts", "body", "size_pt"),
        "h1_font":        _dig(raw, "format", "fonts", "h1", "family"),
        "h1_pt":          _dig(raw, "format", "fonts", "h1", "size_pt"),
        "clause_font":    _dig(raw, "format", "fonts", "h2", "family"),
        "title_pt":       _dig(raw, "format", "fonts", "main_title", "size_pt"),
        "subtitle_pt":    _dig(raw, "format", "fonts", "h3", "size_pt"),
        "indent_cm":      _dig(raw, "format", "paragraph", "first_line_indent_cm"),
        "line_pt":        _dig(raw, "format", "paragraph", "line_spacing_value"),
        "page_number":    _dig(raw, "format", "page", "page_number_position"),
        # 致辞/讲话（speech）题注区与字数注记
        "byline_font":    _dig(raw, "format", "fonts", "byline", "family"),
        "byline_pt":      _dig(raw, "format", "fonts", "byline", "size_pt"),
        "date_font":      _dig(raw, "format", "fonts", "date_line", "family"),
        "date_pt":        _dig(raw, "format", "fonts", "date_line", "size_pt"),
        "sig_font":       _dig(raw, "format", "fonts", "word_count_note", "family"),
        "sig_pt":         _dig(raw, "format", "fonts", "word_count_note", "size_pt"),
    }
    for k, v in font_map.items():
        if v not in (None, "", 0):
            prof[k] = v
    if all(k in m for k in ("top", "bottom", "left", "right")):
        prof["margin"] = (m["top"], m["bottom"], m["left"], m["right"])
    # 副标题字号：档案里 h3 是三级层次（仿宋），副标题应取 h1 或固定 18
    sub = _dig(raw, "format", "fonts", "h1", "size_pt")
    if sub:
        prof["subtitle_pt"] = max(18, int(sub))
    prof["_source"] = str(yaml_path)
    return prof


def load_profile(leader: str, assets_dir: Path | None = None,
                 explicit: str | None = None, doc_type: str = "report") -> dict:
    key = resolve_leader(leader)
    if explicit:
        p = Path(explicit)
        if not p.is_file():
            raise SystemExit(f"[error] 找不到风格档案: {p}")
        prof = profile_from_yaml(p)
        print(f"[info] 版式取自显式档案 {p}")
        return prof
    if assets_dir and key:
        genre_suffix = {"report": "总结", "speech": "致辞", "baogao": "呈报"}.get(doc_type)
        cands = []
        if genre_suffix:
            # 文种专属档案优先（profile_<领导>_<文种>_v*.yaml，文种=总结/致辞/呈报）
            cands = sorted(assets_dir.glob(f"profile_{key}_{genre_suffix}_v*.yaml"))
        if not cands:
            cands = sorted(assets_dir.glob(f"profile_{key}_v*.yaml"))
        if cands:
            prof = profile_from_yaml(cands[-1])
            if doc_type == "speech":
                prof.setdefault("page_number", "none")
            print(f"[info] 版式取自档案 {cands[-1].name}"
                  + (f"（doc-type: {doc_type}）" if doc_type != "report" else ""))
            return prof
    if key in FALLBACK_PROFILES:
        prof = dict(FALLBACK_PROFILES[key])
        prof["_source"] = "builtin-fallback"
        if doc_type == "speech":
            prof.update(SPEECH_DEFAULTS)
        return prof
    prof = dict(DEFAULT_PROFILE)
    prof["_source"] = "default"
    if doc_type == "speech":
        prof.update(SPEECH_DEFAULTS)
    return prof


# ============================================================
# 2. 文本与版式工具
# ============================================================
def is_cjk(ch: str) -> bool:
    return bool(ch) and "\u4e00" <= ch <= "\u9fff"


def normalize_cn_punct(text: str) -> str:
    """中文语境内半角标点 -> 全角；保留 9:00 / 7.9% / 0755-1234 等用法"""
    pairs = {",": "，", ";": "；", "?": "？", "!": "！", ":": "："}
    out = []
    for i, ch in enumerate(text):
        if ch in pairs:
            prev = text[i - 1] if i > 0 else ""
            nxt = text[i + 1] if i + 1 < len(text) else ""
            if is_cjk(prev) or is_cjk(nxt) or (not prev and not nxt):
                out.append(pairs[ch])
                continue
        out.append(ch)
    return "".join(out)


def is_salutation(text: str, max_len: int = 60) -> bool:
    """称呼语判定：必须以冒号结尾、长度受限、命中话题白名单、且不含阻断词"""
    t = text.strip()
    if not t or len(t) > max_len:
        return False
    if not re.search(r"[：:]\s*$", t):
        return False
    head = re.sub(r"[：:]\s*$", "", t)
    if not head:
        return False
    if SALUTATION_BLOCK.search(head):
        return False
    if SALUTATION_HINT.search(head):
        return True
    # 纯姓名列举：「利民、尹杰、伟中：」
    if NAME_LIST_RE.fullmatch(head):
        return True
    # 主送机关：「水利部：」「生态环境部：」「XX市人民政府：」（呈报类公文）
    if MASTER_ORGAN_RE.fullmatch(t):
        return True
    return False


def set_run_font(run, han=FONT_BODY, latin=FONT_LATIN, size=Pt(16), bold=False, color=None):
    run.font.name = latin
    run.font.size = size
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), han)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)


def add_para(doc, text, *, han=FONT_BODY, size=Pt(16), indent_cm=0.0, align=None,
             bold=False, italic=False, line_pt=28, hanging_cm=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = Pt(line_pt)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    if hanging_cm is not None:
        pf.left_indent = Cm(hanging_cm)
        pf.first_line_indent = Cm(-hanging_cm / 2)
    else:
        pf.first_line_indent = Cm(indent_cm)
    run = p.add_run(text)
    set_run_font(run, han=han, size=size, bold=bold)
    run.font.italic = italic
    return p


INLINE_ITALIC_RE = re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)")


def inline_tokens(line: str):
    """把一行拆成 [(起始偏移, 文本, 是否加粗), ...]，用于按偏移决定字体"""
    toks, last = [], 0
    for m in re.finditer(r"\*\*(.+?)\*\*", line):
        if m.start() > last:
            toks.append((last, line[last:m.start()], False))
        toks.append((m.start(), m.group(1), True))
        last = m.end()
    if last < len(line):
        toks.append((last, line[last:], False))
    return toks or [(0, line, False)]


def emit_inline(p, text, *, han, size, bold=False, emit=lambda x: x):
    """把一段文本按 **加粗** / *斜体* / `等宽` 拆成多个 run 写入段落 p"""
    for i, part in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if not part:
            continue
        b = bold or (i % 2 == 1)
        for seg, italic in _split_italic(part):
            for code_seg, mono in _split_code(seg):
                if not code_seg:
                    continue
                run = p.add_run(emit(code_seg))
                set_run_font(run, han=han, size=size, bold=b)
                run.font.italic = italic
                if mono:
                    run.font.name = "Consolas"
    return p


def _new_para(doc, *, indent_cm=0.0, line_pt=28, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.first_line_indent = Cm(indent_cm)
    pf.line_spacing = Pt(line_pt)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    return p


def add_rich_para(doc, line, *, han=FONT_BODY, size=Pt(16), indent_cm=0.0, line_pt=28,
                  align=None, force_bold=False):
    """普通段落，支持 **加粗**、*斜体*、`等宽` 三类行内富文本"""
    p = _new_para(doc, indent_cm=indent_cm, line_pt=line_pt, align=align)
    emit_inline(p, line, han=han, size=size, bold=force_bold)
    return p


def add_clause_lead_para(doc, line, *, lead_end, clause_font=FONT_KAI,
                         body_font=FONT_BODY, size=Pt(16), indent_cm=0.0,
                         line_pt=28, emit=lambda x: x):
    """条款引语段落：引语段（含行内加粗）用楷体，其余用仿宋正文"""
    p = _new_para(doc, indent_cm=indent_cm, line_pt=line_pt)
    for off, seg, bold in inline_tokens(line):
        if off < lead_end < off + len(seg):
            # token 跨越引语边界（纯文本整段一个 token 的常见情形）：
            # 按边界拆成 楷体引语 + 仿宋正文 两段，否则整段都会落成楷体
            head, tail = seg[:lead_end - off], seg[lead_end - off:]
            emit_inline(p, head, han=clause_font, size=size, bold=bold, emit=emit)
            emit_inline(p, tail, han=body_font, size=size, bold=bold, emit=emit)
            continue
        han = clause_font if off < lead_end else body_font
        emit_inline(p, seg, han=han, size=size, bold=bold, emit=emit)
    return p


def _split_italic(part: str):
    out, last = [], 0
    for m in INLINE_ITALIC_RE.finditer(part):
        if m.start() > last:
            out.append((part[last:m.start()], False))
        out.append((m.group(1), True))
        last = m.end()
    if last < len(part):
        out.append((part[last:], False))
    return out or [(part, False)]


def _split_code(part: str):
    out, last = [], 0
    for m in re.finditer(r"`([^`]+)`", part):
        if m.start() > last:
            out.append((part[last:m.start()], False))
        out.append((m.group(1), True))
        last = m.end()
    if last < len(part):
        out.append((part[last:], False))
    return out or [(part, False)]


def add_page_number_footer(section, han=FONT_BODY, size=Pt(14)):
    """页脚居中页码域（档案 page_number_position: bottom-center）"""
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    run = p.add_run()
    set_run_font(run, han=han, size=size)
    for tag, attrs, text in (("w:fldChar", {"w:fldCharType": "begin"}, None),
                             ("w:instrText", {"xml:space": "preserve"}, " PAGE "),
                             ("w:fldChar", {"w:fldCharType": "end"}, None)):
        el = OxmlElement(tag)
        for k, v in attrs.items():
            el.set(qn(k), v)
        if text:
            el.text = text
        run._r.append(el)
    return p


def add_three_line_table(doc, rows, *, body_font=FONT_BODY, head_font=FONT_HEI,
                         size_pt=14, line_pt=22):
    """Markdown 表格 -> docx 三线表（表头上下线 + 表尾底线，无竖线）"""
    n_cols = max(len(r) for r in rows)
    table = doc.add_table(rows=0, cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for ri, row in enumerate(rows):
        cells = table.add_row().cells
        for ci in range(n_cols):
            txt = row[ci] if ci < len(row) else ""
            cell = cells[ci]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = p.paragraph_format
            pf.line_spacing = Pt(line_pt)
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.first_line_indent = Cm(0)
            run = p.add_run(txt)
            set_run_font(run, han=(head_font if ri == 0 else body_font),
                         size=Pt(size_pt), bold=(ri == 0))
    # 清掉全部默认框线
    tblPr = table._tbl.tblPr
    for el in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(el)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "none")
        e.set(qn("w:sz"), "0")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        borders.append(e)
    _insert_ordered(tblPr, borders, TBLPR_ORDER)
    # 三线：表头上下 + 表尾下
    def row_edge(row, edge, sz):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for old in tcPr.findall(qn("w:tcBorders")):
                tcPr.remove(old)
            b = OxmlElement("w:tcBorders")
            e = OxmlElement(f"w:{edge}")
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), str(sz))
            e.set(qn("w:space"), "0")
            e.set(qn("w:color"), "000000")
            b.append(e)
            _insert_ordered(tcPr, b, TCPR_ORDER)

    if table.rows:
        row_edge(table.rows[0], "top", 12)
        row_edge(table.rows[0], "bottom", 6)
        row_edge(table.rows[-1], "bottom", 12)
    # 表后留一空行，避免与下段粘连
    doc.add_paragraph()
    return table


def parse_md_table(lines, start):
    """从 lines[start] 起解析 Markdown 管道表格，返回 (rows, next_index)"""
    def cells(l):
        s = l.strip().strip("|")
        return [c.strip() for c in re.split(r"(?<!\\)\|", s)]

    header = cells(lines[start])
    if start + 1 >= len(lines) or not TABLE_SEP_RE.match(lines[start + 1]):
        return None, start + 1
    rows = [header]
    i = start + 2
    while i < len(lines) and lines[i].strip().startswith("|"):
        rows.append(cells(lines[i]))
        i += 1
    return rows, i


# ============================================================
# 3. 主转换
# ============================================================
def build(md_path: Path, out_path: Path, leader: str = "", *,
          profile_path: str | None = None, strict_leader: bool = False,
          clause_scope: str = "lead", h3_bold: bool = False,
          normalize_punct: bool = False, page_number: bool | None = None,
          assets_dir: Path | None = None, quiet: bool = False,
          pdf: bool = False, pdf_path: str | Path | None = None,
          doc_type: str = "report") -> dict:
    key = resolve_leader(leader)
    speech = (doc_type == "speech")
    unknown = bool(leader) and key not in FALLBACK_PROFILES
    if strict_leader and (unknown or not leader):
        raise SystemExit(
            f"[error] --strict-leader：未识别的领导 '{leader}'。"
            f"可用值：{', '.join(FALLBACK_PROFILES)}（或别名 {', '.join(ALIASES)}），"
            f"也可用 --profile 直接指定档案 YAML。")
    prof = load_profile(leader, assets_dir=assets_dir, explicit=profile_path,
                        doc_type=doc_type)
    if speech:                       # speech 兜底（档案缺题注字段时）
        for k, v in SPEECH_DEFAULTS.items():
            prof.setdefault(k, v)
    if unknown and not profile_path:
        print(f"[warn] 未识别的领导 '{leader}'，回退到默认版式（三号 16pt / 缩进 1.13cm）")

    body_pt = Pt(prof["body_pt"])
    indent = prof["indent_cm"]
    line_pt = prof.get("line_pt", 28)
    top, bottom, left, right = prof["margin"]

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = FONT_LATIN
    style.font.size = body_pt
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), prof["body_font"])
    rfonts.set(qn("w:ascii"), FONT_LATIN)
    rfonts.set(qn("w:hAnsi"), FONT_LATIN)

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)   # A4（原缺，导致默认 Letter）
    sec.top_margin, sec.bottom_margin = Cm(top / 10), Cm(bottom / 10)
    sec.left_margin, sec.right_margin = Cm(left / 10), Cm(right / 10)

    want_pn = (prof.get("page_number") == "bottom-center") if page_number is None else page_number
    if want_pn:
        add_page_number_footer(sec, han=prof["body_font"], size=Pt(max(10, prof["body_pt"] - 4)))

    lines = md_path.read_text(encoding="utf-8").splitlines()
    stats = {"paragraph": 0, "table": 0, "table_rows": 0, "list": 0,
             "h1": 0, "h2": 0, "h3": 0, "h4": 0, "clause": 0,
             "salutation": 0, "signature": 0, "skipped": 0, "in_table": False,
             "byline": 0, "date_line": 0, "word_note": 0}
    speech_chars = 0          # speech 模式：称呼+正文累计字数（自动补字数注记用）

    def emit(text):
        return normalize_cn_punct(text) if normalize_punct else text

    i = 0
    while i < len(lines):
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()
        i += 1

        if not stripped or stripped.startswith(">") or stripped == "---":
            stats["skipped"] += 1
            continue

        # ---- 表格 ----
        if stripped.startswith("|"):
            rows, nxt = parse_md_table(lines, i - 1)
            if rows:
                add_three_line_table(
                    doc, rows, body_font=prof["body_font"], head_font=FONT_HEI,
                    size_pt=max(10, prof["body_pt"] - 4), line_pt=max(16, line_pt - 6))
                stats["table"] += 1
                stats["table_rows"] += sum(len(r) for r in rows)
                i = nxt
                continue

        # ---- 标题层级 ----
        if stripped.startswith("#### "):
            add_rich_para(doc, emit(stripped[5:]), han=prof["clause_font"], size=body_pt,
                          indent_cm=indent, line_pt=line_pt, force_bold=True)
            stats["h4"] += 1
            continue
        if stripped.startswith("### "):
            add_rich_para(doc, emit(stripped[4:]), han=prof["h1_font"], size=Pt(prof["h1_pt"]),
                          indent_cm=indent, line_pt=line_pt)
            stats["h1"] += 1
            continue
        if stripped.startswith("## "):
            text2 = emit(stripped[3:])
            if speech:
                # 致辞题注区：日期行（（YYYY年M月D日））仿宋居中 / 署名行（职务+姓名）楷体居中
                if DATE_LINE_RE.match(text2):
                    add_para(doc, text2, han=prof.get("date_font", FONT_BODY),
                             size=Pt(prof.get("date_pt", 16)),
                             align=WD_ALIGN_PARAGRAPH.CENTER, line_pt=line_pt)
                    stats["date_line"] += 1
                else:
                    add_para(doc, text2, han=prof.get("byline_font", FONT_KAI),
                             size=Pt(prof.get("byline_pt", 16)),
                             align=WD_ALIGN_PARAGRAPH.CENTER, line_pt=line_pt)
                    stats["byline"] += 1
                continue
            add_para(doc, text2, han=FONT_HEI, size=Pt(prof["subtitle_pt"]),
                     align=WD_ALIGN_PARAGRAPH.CENTER, line_pt=line_pt)
            stats["h2"] += 1
            continue
        if stripped.startswith("# "):
            add_para(doc, emit(stripped[2:]), han=FONT_TITLE, size=Pt(prof["title_pt"]),
                     align=WD_ALIGN_PARAGRAPH.CENTER, line_pt=max(line_pt, 34))
            stats["h3"] += 1
            continue

        # ---- 落款（**整行**）----
        m_sig = re.fullmatch(r"\*\*(.+)\*\*", stripped)
        if m_sig and len(stripped) < 80:
            sig_text = emit(m_sig.group(1))
            if speech and WORDCOUNT_NOTE_RE.match(sig_text):
                # 致辞文末字数注记「（共N字，约M分钟）」：楷体 16pt 右对齐
                add_para(doc, sig_text, han=prof.get("sig_font", FONT_KAI),
                         size=Pt(prof.get("sig_pt", 16)),
                         align=WD_ALIGN_PARAGRAPH.RIGHT, line_pt=line_pt)
                stats["word_note"] += 1
                continue
            if speech:
                print(f"[warn] speech 模式署名应在标题区（## 职务+姓名），"
                      f"文末落款已按楷体注记渲染：{sig_text[:30]}")
                add_para(doc, sig_text, han=prof.get("sig_font", FONT_KAI),
                         size=Pt(prof.get("sig_pt", 16)),
                         align=WD_ALIGN_PARAGRAPH.RIGHT, line_pt=line_pt)
                stats["signature"] += 1
                continue
            add_para(doc, sig_text, han=prof["body_font"], size=body_pt,
                     align=WD_ALIGN_PARAGRAPH.RIGHT, line_pt=line_pt)
            stats["signature"] += 1
            continue

        # ---- 称呼语（顶格，不缩进）----
        if is_salutation(stripped):
            add_rich_para(doc, emit(stripped), han=prof["body_font"], size=body_pt,
                          indent_cm=0.0, line_pt=line_pt)
            stats["salutation"] += 1
            if speech:
                speech_chars += len(stripped)
            continue

        # ---- 无序列表 ----
        m_ul = LIST_UL_RE.match(line)
        if m_ul:
            add_rich_para(doc, emit(m_ul.group(1).strip()), han=prof["body_font"],
                          size=body_pt, indent_cm=indent, line_pt=line_pt)
            stats["list"] += 1
            if speech:
                speech_chars += len(m_ul.group(1))
            continue

        # ---- 有序列表 / 三级层次 ----
        m_ol = LIST_OL_RE.match(line)
        if m_ol:
            add_rich_para(doc, emit(f"{m_ol.group(1)}.{m_ol.group(2).strip()}"),
                          han=prof["body_font"], size=body_pt, indent_cm=indent,
                          line_pt=line_pt, force_bold=h3_bold)
            stats["list"] += 1
            if speech:
                speech_chars += len(m_ol.group(1)) + len(m_ol.group(2))
            continue

        # ---- 一级标题「一、」（baogao 呈报模式：黑体独立短行）----
        if doc_type == "baogao" and LEVEL1_RE.match(stripped) and len(stripped) <= 50 \
                and not re.search(r"[。；]", stripped):
            add_rich_para(doc, emit(stripped), han=prof.get("h1_font", FONT_HEI),
                          size=body_pt, indent_cm=indent, line_pt=line_pt)
            stats["h1"] += 1
            continue

        # ---- 条款（（一）…）----
        if CLAUSE_RE.match(stripped):
            stats["clause"] += 1
            if speech:
                speech_chars += len(stripped)
            m = CLAUSE_LEAD_INLINE_RE.match(stripped) if clause_scope == "lead" else None
            if m and m.end() < len(stripped):
                add_clause_lead_para(doc, stripped, lead_end=m.end(),
                                     clause_font=prof["clause_font"],
                                     body_font=prof["body_font"], size=body_pt,
                                     indent_cm=indent, line_pt=line_pt, emit=emit)
            else:
                add_rich_para(doc, emit(stripped), han=prof["clause_font"], size=body_pt,
                              indent_cm=indent, line_pt=line_pt)
            continue

        # ---- 普通正文 ----
        add_rich_para(doc, emit(stripped), han=prof["body_font"], size=body_pt,
                      indent_cm=indent, line_pt=line_pt)
        if speech:
            speech_chars += len(stripped)

    # ---- speech 模式：文末自动补「（共N字，约M分钟）」注记（约 200 字/分钟）----
    if speech and stats["word_note"] == 0 and speech_chars > 0:
        minutes = max(1, round(speech_chars / 200))
        add_para(doc, f"（共{speech_chars}字，约{minutes}分钟）",
                 han=prof.get("sig_font", FONT_KAI), size=Pt(prof.get("sig_pt", 16)),
                 align=WD_ALIGN_PARAGRAPH.RIGHT, line_pt=line_pt)
        stats["word_note"] += 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    stats["paragraph"] = len([p for p in doc.paragraphs if p.text.strip()])
    stats["chars"] = sum(len(p.text) for p in doc.paragraphs)
    stats["profile_source"] = prof.get("_source", "?")
    stats["pdf"] = False
    if pdf:
        try:
            from pdf_export import docx_to_pdf       # 懒加载：仅 Windows 需要 pywin32
            stats["pdf"] = str(docx_to_pdf(out_path, pdf_path, quiet=quiet))
        except Exception as e:                        # noqa: BLE001
            stats["pdf_error"] = f"{type(e).__name__}: {e}"
            print(f"[warn] PDF 导出失败（docx 已正常生成）：{e}")
    if not quiet:
        print(f"OK  -> {out_path}")
        print(f"    文种: {doc_type} | 版式: {leader or '默认'} | 正文 {prof['body_pt']}pt "
              f"| 缩进 {indent}cm | 边距 {top}/{bottom}/{left}/{right}mm "
              f"| 页码 {'开' if want_pn else '关'}")
        extra = (f" | 署名行 {stats['byline']} | 日期行 {stats['date_line']} "
                 f"| 字数注记 {stats['word_note']}") if speech else ""
        print(f"    段落 {stats['paragraph']} | 表格 {stats['table']}({stats['table_rows']}格) "
              f"| 一级标题 {stats['h1']} | 条款 {stats['clause']} | 列表 {stats['list']} "
              f"| 字数 {stats['chars']}{extra}")
        if pdf and stats.get("pdf"):
            print(f"    PDF  {stats['pdf']}")
    return stats


def main():
    ap = argparse.ArgumentParser(description="Markdown 报告 -> 集团公文档式 docx")
    ap.add_argument("--input", required=True, help="Markdown 源文件")
    ap.add_argument("--output", required=True, help="输出 docx 路径")
    ap.add_argument("--leader", default="", help="领导姓名（龚利民 / 方琳），决定版式")
    ap.add_argument("--doc-type", default="report",
                    help="文种：report=总结/汇报（默认）；speech=致辞/讲话（题注式，"
                         "亦接受 致辞/讲话/发言/总结/报告 等中文别名）")
    ap.add_argument("--profile", default=None, help="显式指定风格档案 YAML（优先于内置）")
    ap.add_argument("--strict-leader", action="store_true", help="未识别领导时报错而非回退")
    ap.add_argument("--clause-scope", choices=["lead", "whole"], default="lead",
                    help="条款楷体范围：lead=仅引语（贴合两位领导实测档案，默认）；whole=整段")
    ap.add_argument("--h3-bold", action="store_true", help="三级层次（1.）加粗")
    ap.add_argument("--no-page-number", action="store_true", help="不生成页脚页码")
    ap.add_argument("--normalize-punct", action="store_true", help="中文语境半角标点转全角")
    ap.add_argument("--pdf", action="store_true",
                    help="同时导出 PDF（需 Windows + Microsoft Word）")
    ap.add_argument("--pdf-out", default=None, help="PDF 输出路径（默认与 docx 同名）")
    args = ap.parse_args()

    md = Path(args.input)
    if not md.is_file():
        raise SystemExit(f"[error] 找不到输入文件: {md}")
    doc_type = DOC_TYPE_ALIASES.get(args.doc_type.strip().lower(), "report")
    assets = Path(__file__).resolve().parent.parent / "assets"
    build(md, Path(args.output), args.leader,
          profile_path=args.profile, strict_leader=args.strict_leader,
          clause_scope=args.clause_scope, h3_bold=args.h3_bold,
          normalize_punct=args.normalize_punct,
          page_number=False if args.no_page_number else None,
          assets_dir=assets,
          pdf=args.pdf, pdf_path=args.pdf_out,
          doc_type=doc_type)


if __name__ == "__main__":
    main()
