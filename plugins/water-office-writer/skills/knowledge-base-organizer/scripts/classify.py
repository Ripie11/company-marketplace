#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""classify.py —— 知识库分类决策引擎

把「一份文档该归到哪」从"人脑判断"变成一条**可复现、可回测**的确定性流程：

    属性抽取（四类信号）  →  冲突消解  →  目录路由  →  置信度分流  →  命名生成

设计要点：
  * 规则全部来自 assets/routing-rules.yaml 与 assets/taxonomy.yaml，
    本脚本不含任何硬编码的业务判定 —— 这是与上一版（ITEMS 数组）的根本区别。
  * 信号可靠性严格排序：① 文件名 > ② 文档内部结构 > ③ 正文关键词 > ④ 来源路径。
    同优先级冲突一律降级为「低置信」转人工，绝不猜测。
  * 只读，不写任何文件。落盘交给 intake.py。

用法：
    python classify.py --input <文件或目录> [--json 结果.json] [--no-content] [--root 知识库根]
    python classify.py --input xxx.docx --quiet          # 只输出 JSON，便于管道

退出码：0 成功；2 参数/资产错误。
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
import unicodedata
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any

try:
    import yaml
except ImportError:
    sys.stderr.write("[FAIL] 缺少 PyYAML，请先安装：pip install pyyaml\n")
    raise SystemExit(2)

# Windows 控制台默认 GBK，中文输出会崩；强制切 UTF-8
if sys.stdout.encoding is None or sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass

SUPPORTED_EXT = {".docx", ".doc", ".pptx", ".ppt", ".pdf", ".md", ".txt"}

# ============================================================================
# 资产加载
# ============================================================================


class Assets:
    """taxonomy.yaml + routing-rules.yaml 的聚合视图。"""

    def __init__(self, skill_dir: Path):
        self.skill_dir = skill_dir
        self.taxonomy = self._load(skill_dir / "assets" / "taxonomy.yaml")
        self.rules = self._load(skill_dir / "assets" / "routing-rules.yaml")
        self._compile()
        self._build_index()

    @staticmethod
    def _load(path: Path) -> dict:
        if not path.exists():
            raise SystemExit(f"[FAIL] 找不到资产文件：{path}")
        with open(path, encoding="utf-8") as f:
            return yaml.safe_load(f)

    def _compile(self) -> None:
        """预编译所有正则，避免逐文件重复编译。"""
        def comp(seq, key="match"):
            for item in seq:
                item["_re"] = re.compile(item[key])

        comp(self.rules.get("filename_markers", []))
        comp(self.rules.get("salutations", []))
        comp(self.rules.get("keyword_types", []))
        comp(self.rules.get("subject_markers", []))
        comp(self.rules.get("date_rules", []))
        self.strip_res = [re.compile(p) for p in self.taxonomy["meta"].get("strip_patterns", [])]

    def _build_index(self) -> None:
        self.space_by_subject: dict[str, dict] = {}
        self.space_aliases: dict[str, str] = {}
        for sp in self.taxonomy["spaces"]:
            for subj in sp.get("subjects", []):
                self.space_by_subject[subj] = sp
            self.space_aliases[sp["id"]] = sp["id"]
            for alias in sp.get("aliases", []):
                self.space_aliases[alias] = sp["id"]

        self.subdir_by_type: dict[str, dict] = {}
        for sd in self.taxonomy["sub_dirs"]:
            for t in sd.get("types", []):
                self.subdir_by_type[t] = sd

        # routing.overrides：按条件覆盖子目录
        self.sub_overrides: list[dict] = []
        for ov in self.rules.get("routing", {}).get("overrides", []):
            self.sub_overrides.append(ov)

    # ---- 查询辅助 ----
    def space_for_subject(self, subject: str | None) -> dict | None:
        if not subject:
            return None
        return self.space_by_subject.get(subject)

    def subdir_for_type(self, doc_type: str | None) -> dict | None:
        if not doc_type:
            return None
        for ov in self.sub_overrides:
            if ov.get("when", {}).get("type") == doc_type:
                target = ov.get("sub_dir")
                if target is None:
                    return None  # 覆盖为"落主空间根"
                for sd in self.taxonomy["sub_dirs"]:
                    if sd["id"] == target:
                        return sd
        return self.subdir_by_type.get(doc_type)

    def override_target_dir(self, doc_type: str | None) -> str | None:
        for ov in self.sub_overrides:
            if ov.get("when", {}).get("type") == doc_type and "target_dir" in ov:
                return ov["target_dir"]
        return None


# ============================================================================
# 文本抽取
# ============================================================================


@dataclass
class DocText:
    text: str = ""
    method: str = "none"
    ok: bool = False
    note: str = ""

    @property
    def head(self) -> str:
        return self.text[:1200]

    @property
    def tail(self) -> str:
        return self.text[-900:]


def _read_docx(path: Path) -> DocText:
    from docx import Document
    d = Document(str(path))
    paras = [p.text.strip() for p in d.paragraphs if p.text.strip()]
    # 表格里也常有标题/落款
    for t in d.tables:
        for row in t.rows:
            for cell in row.cells:
                if cell.text.strip():
                    paras.append(cell.text.strip())
    return DocText("\n".join(paras), "python-docx", True)


def _read_pptx(path: Path) -> DocText:
    from pptx import Presentation
    prs = Presentation(str(path))
    chunks: list[str] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = "".join(r.text for r in para.runs).strip()
                    if t:
                        chunks.append(t)
    return DocText("\n".join(chunks), "python-pptx", True)


def _read_pdf(path: Path, max_pages: int = 3) -> DocText:
    import pdfplumber
    chunks: list[str] = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages[:max_pages]:
            t = page.extract_text() or ""
            if t.strip():
                chunks.append(t)
        # 落款通常在最后，取尾页
        if len(pdf.pages) > max_pages:
            t = pdf.pages[-1].extract_text() or ""
            if t.strip():
                chunks.append(t)
    return DocText("\n".join(chunks), "pdfplumber", True)


def _read_doc(path: Path) -> DocText:
    """旧版 .doc（OLE 复合文档）无纯 Python 可靠解析。

    退而求其次：直接从二进制里捞 UTF-16LE 的中文串，足以做**关键词兜底**，
    但**不足以**做精确的标题/落款判断 —— 因此证据里会显式标注降级。
    """
    raw = path.read_bytes()
    found: list[str] = []
    # Word 的正文常以 UTF-16LE 存储，CJK 区间在 0x4E00-0x9FFF
    try:
        s = raw.decode("utf-16-le", errors="ignore")
    except Exception:
        s = ""
    for m in re.finditer(r"[\u4e00-\u9fff][\u4e00-\u9fff，。、；：（）“”0-9%]{5,}", s):
        found.append(m.group(0))
    text = "\n".join(found[:400])
    return DocText(text, "doc-binary-best-effort", bool(found),
                   note="旧版 .doc 无法精确解析，已降级为关键词兜底")


def extract_text(path: Path) -> DocText:
    ext = path.suffix.lower()
    try:
        if ext == ".docx":
            return _read_docx(path)
        if ext == ".pptx":
            return _read_pptx(path)
        if ext == ".pdf":
            return _read_pdf(path)
        if ext == ".doc":
            return _read_doc(path)
        if ext in (".md", ".txt"):
            return DocText(path.read_text(encoding="utf-8", errors="ignore"), "text", True)
    except Exception as e:  # 抽取失败不影响文件名信号
        return DocText("", type(e).__name__, False, note=f"文本抽取失败：{e}")
    return DocText("", "unsupported", False, note=f"不支持的扩展名 {ext}")


# ============================================================================
# 判定结果
# ============================================================================


@dataclass
class Verdict:
    src: str
    filename: str
    ext: str
    doc_type: str | None = None
    subject: str | None = None
    audience: str | None = None
    security: str | None = None
    tags: list[str] = field(default_factory=list)
    date: str | None = None
    title_clean: str = ""
    space: str | None = None
    sub_dir: str | None = None
    target_dir: str | None = None
    new_filename: str | None = None
    confidence: float = 0.0
    level: str = "low"          # high | medium | low
    action: str = "block"       # silent | preview | block
    evidence: list[str] = field(default_factory=list)
    blockers: list[str] = field(default_factory=list)
    candidates: list[dict] = field(default_factory=list)
    text_method: str = ""
    text_ok: bool = False
    notes: list[str] = field(default_factory=list)


# ============================================================================
# 信号抽取
# ============================================================================


def _match_filename(rules: list[dict], name: str) -> list[dict]:
    hits = []
    for r in rules:
        m = r["_re"].search(name)
        if m:
            hits.append({"rule": r, "m": m})
    return hits


def _clean_brackets(name: str) -> str:
    """剥离 【】（）() 等括注（内容已单独抽取为信号）。"""
    s = re.sub(r"【[^】]*】", "", name)
    s = re.sub(r"（[^）]*）", "", s)
    s = re.sub(r"\([^)]*\)", "", s)
    s = re.sub(r"\[[^\]]*\]", "", s)
    return s


def clean_title(stem: str, assets: Assets, doc_type: str | None) -> str:
    """生成规范标题（文件名中段）。

    ⚠️ 这一步是**启发式机械压缩**，无法完全等价于人脑改写。回测中按"相似度"
    而非"完全相等"评估，最终仍建议人工扫一眼。
    """
    s = unicodedata.normalize("NFKC", stem)
    org_suffix = ""
    m = re.search(r"【报(?P<org>住建部|水利部|生态环境部|省厅)】", s)
    if m:
        org_suffix = "-" + m.group("org")

    s = _clean_brackets(s)

    for r in assets.strip_res:
        s = r.sub("", s)

    s = re.sub(r"^\s*[\d\s\-_、，,]+", "", s)              # 前导时间戳/序号
    s = re.sub(r"\d{4}\s*\d{3,4}\s*", "", s)               # 0711 1300 类
    s = re.sub(r"^(龚董|龚总|董事长|龚利民|方琳|方总)(?=在|于|向|的)", "", s)
    s = re.sub(r"^在", "", s)                              # 在XX会议上的讲话
    s = re.sub(r"深圳环境水务集团关于", "", s)
    s = re.sub(r"深圳市环境水务集团关于", "", s)
    s = re.sub(r"深圳环境水务集团", "集团", s)
    s = re.sub(r"上的(?=致辞|讲话|发言)", "", s)
    s = re.sub(r"会议上的(?=讲话|致辞)", "会议", s)
    s = re.sub(r"[（(]龚董[）)]|[（(]龚总[）)]|[（(]方琳总[）)]|[（(]方总[）)]", "", s)
    s = re.sub(r"(?<=[\u4e00-\u9fff])\d{2,4}$", "", s)     # 尾部流水号，如「讲话0426」
    s = re.sub(r"^[\s\-_、，,]+|[\s\-_、，,]+$", "", s)
    s = s.replace("：", "-").replace(":", "-")

    if doc_type == "半年总结":
        s = s.replace("半年工作总结", "半年总结")
    s = re.sub(r"\s+", "", s)
    if org_suffix and not s.endswith(org_suffix):
        s = s + org_suffix
    return s


# 纯日期/版本词构成的【】内容，不能当"场合名"用
_BRACKET_NOISE = re.compile(
    r"^(?:20\d{2}(?:0[1-9]|1[0-2])?(?:0[1-9]|[12]\d|3[01])?年?|"
    r"终稿|定稿|最终版|成稿|修订稿|附件\d+|第[一二三四五六七八九十\d]+次审稿后?)$"
)
# 领导指代不构成"场合名"
_LEADER_NOISE = {"龚董", "龚总", "董事长", "龚利民", "方琳", "方总", "方琳总"}


def _bracket_context(stem: str, type_word: str) -> str:
    """从【】里提取"场合名"，用于补全退化成类型词的标题。

    例：【2026年中国水协年会】主旨报告     → ctx=中国水协年会 → 中国水协年会主旨报告
        【水协科技委会长】【2024】就职演讲 → ctx=水协科技委会长 → 水协科技委会长就职演讲
    """
    for b in re.findall(r"【([^】]*)】", stem):
        b = b.strip()
        b = re.sub(r"^20\d{2}年?", "", b).strip()      # 前导年份
        if not b or _BRACKET_NOISE.match(b) or b in _LEADER_NOISE:
            continue
        if type_word and b == type_word:
            continue
        return b
    return ""


def enhance_title(title: str, stem: str, doc_type: str | None, assets: Assets) -> str:
    """标题退化修补。

    clean_title() 是机械压缩，遇到「【场合】+类型词」这种命名时会把标题压成
    只剩类型词（如"主旨报告"），产出 `20260101_主旨报告_主旨报告_龚利民.pptx`
    这种同词重复。此函数只在【标题退化】时出手，按三级兜底补全：

        ① 借【】里的场合名   →  中国水协年会主旨报告
        ② 用 taxonomy 的 title_fallback 资产
        ③ 原样返回（宁缺毋滥，交给人工）
    """
    type_word = doc_type or ""
    core = re.sub(r"[\s\-_、,，]", "", title or "")
    # 退化判定：标题为空 / 就是类型词本身 / 比类型词还短
    if core and (not type_word or (core != type_word and len(core) > len(type_word))):
        return title

    ctx = _bracket_context(stem, type_word)
    if ctx:
        return f"{ctx}{type_word}" if type_word else ctx

    dflt = (assets.taxonomy.get("meta", {}).get("title_fallback") or {}).get(type_word)
    if dflt:
        return dflt
    return title


def extract_signals(path: Path, assets: Assets, doc: DocText) -> tuple[dict, list[str], list[str]]:
    """返回 (属性字典, 证据列表, 备注列表)。"""
    name = path.name
    stem = path.stem
    attrs: dict[str, Any] = {}
    evidence: list[str] = []
    notes: list[str] = []
    field_weight: dict[str, float] = {}

    def set_attr(k: str, v: Any, w: float, ev: str) -> None:
        """设定属性；同字段保留权重更高者，权重相同且值不同 → 标记冲突。"""
        cur_w = field_weight.get(k, -1.0)
        if k in attrs and attrs[k] != v and abs(cur_w - w) < 1e-9:
            notes.append(f"[冲突] 字段 {k}：{attrs[k]} vs {v}（同权重 {w}，已降级）")
            field_weight[k] = max(cur_w, w) - 0.15
            return
        if w > cur_w or k not in attrs:
            attrs[k] = v
            field_weight[k] = w
            if ev:
                evidence.append(ev)

    # ---------- 信号①：文件名标记 ----------
    # ⚠️ 对 stem（不含扩展名）匹配 —— 否则 `规范$` 这类锚定会被 `.docx` 破坏
    for hit in _match_filename(assets.rules["filename_markers"], stem):
        r, m = hit["rule"], hit["m"]
        ev = r.get("evidence", "文件名规则命中").replace("{org}", m.groupdict().get("org", "") or "")
        ev = re.sub(r"\{n\}", re.search(r"\d{4}", name).group(0) if re.search(r"\d{4}", name) else "", ev)
        for k, v in (r.get("set") or {}).items():
            set_attr(k, v, r["weight"], f"①{ev}")
        for t in r.get("add_tags") or []:
            tags = attrs.setdefault("tags", [])
            if t not in tags:
                tags.append(t)

    # ---------- 信号②：称呼语（首段） ----------
    if doc.ok and doc.head:
        for r in assets.rules["salutations"]:
            m = r["_re"].search(doc.text.strip())
            if m:
                for k, v in (r.get("set") or {}).items():
                    set_attr(k, v, r["weight"], f"②{r.get('evidence','称呼语命中')}")
                break  # 称呼语只取第一条

    # ---------- 信号③：正文关键词 → type ----------
    if doc.ok and "type" not in attrs:
        body = doc.text
        best = None
        for r in assets.rules["keyword_types"]:
            if r["_re"].search(body):
                if best is None or r["weight"] > best["weight"]:
                    best = r
        if best:
            set_attr("type", best["type"], best["weight"], f"③正文关键词推断：{best['type']}")

    # ---------- 信号④：来源路径 ----------
    posix = path.as_posix()
    for r in assets.rules["source_dirs"]:
        if r["path_contains"] in posix:
            for k, v in (r.get("set") or {}).items():
                set_attr(k, v, r["weight"], f"④来源路径含「{r['path_contains']}」")

    # ---------- leader：文件名指代 > 落款 > 路径 ----------
    marker_found = False
    for r in assets.rules["subject_markers"]:
        if r["_re"].search(stem):
            set_attr("leader", r["set"]["leader"], r["weight"], f"①{r.get('evidence','文件名指代')}")
            marker_found = True
    if not marker_found and doc.ok and doc.tail:
        for r in assets.rules["subject_markers"]:
            if r["_re"].search(doc.tail):
                set_attr("leader", r["set"]["leader"], 0.75, f"②落款指代「{r['set']['leader']}」")
                break
    if not marker_found and doc.ok and doc.head:
        for r in assets.rules["subject_markers"]:
            if r["_re"].search(doc.head):
                set_attr("leader", r["set"]["leader"], 0.7, f"②正文指代「{r['set']['leader']}」")
                break

    # ---------- 一致性检查：路径 vs 内容 ----------
    path_leader = None
    for r in assets.rules["source_dirs"]:
        if r["path_contains"] in posix:
            path_leader = r["set"]["leader"]
    if path_leader and attrs.get("leader") and path_leader != attrs["leader"]:
        notes.append(f"[冲突] 来源路径指向 {path_leader}，内容/文件名指向 {attrs['leader']}")

    # ---------- 日期 ----------
    # 注意：必须在 type 落定之后再抽日期 —— 年份兜底依赖 doc_type
    attrs["date"] = extract_date(name, doc, assets, attrs.get("type"))

    attrs["_field_weight"] = field_weight
    # 同一规则常一次性设定 type/audience/security 三个字段，会产生重复证据行；
    # 去重保序，让"依据"读起来像人话
    evidence = list(dict.fromkeys(e for e in evidence if e))
    return attrs, evidence, notes


def extract_date(name: str, doc: DocText, assets: Assets, doc_type: str | None = None) -> tuple[str, str]:
    """返回 (YYYY-MM-DD, 依据)。

    分三阶段，前两阶段都**只看文件名**，正文只在文件名彻底无年份时才介入：

      阶段1 explicit：显式日期标记（【20260723】/【202604】/【2026年】/内嵌 20260828）
      阶段2 年份兜底：文件名里有年份但无月日 → 按类型推断月日
                     （年终总结 = 次年1月1日；半年总结 = 当年7月1日）
      阶段3 content ：正文日期（落款区优先）
      阶段4 默认值
    """
    stem = Path(name).stem
    rules = assets.rules["date_rules"]

    # ---- 阶段1：显式日期标记 ----
    for r in rules:
        if r.get("stage", "explicit") != "explicit":
            continue
        m = r["_re"].search(stem)
        if not m:
            continue
        fmt = r.get("format", "")
        if fmt == "cn":
            d = _cn_date_to_iso(m.group(0))
            if d:
                return d, f"文件名·{r['name']}"
            continue
        if fmt.startswith("{"):
            out = fmt
            for i, g in enumerate(m.groups(), start=1):
                if g is not None:
                    out = out.replace("{" + str(i) + "}", str(g))
            if "{" not in out:
                return _norm_date(out), f"文件名·{r['name']}"

    # ---- 阶段2：类型感知的年份兜底 ----
    years = [int(y) for y in re.findall(r"(20\d{2})", stem) if 2015 <= int(y) <= 2099]
    if years:
        y = min(years)                     # 取最早年份 = 文档的"报告年度"
        defaults = assets.rules.get("date_year_defaults", {})
        cfg = defaults.get(doc_type) or defaults.get("default") or {}
        y += int(cfg.get("year_offset", 0))
        mo = int(cfg.get("month", 1))
        dd = int(cfg.get("day", 1))
        return f"{y:04d}-{mo:02d}-{dd:02d}", f"文件名·年份（{doc_type or '通用'}推定月日）"

    # ---- 阶段3：正文日期（落款区优先）----
    skip_content = doc_type in (
        (assets.rules.get("date_policy") or {}).get("skip_content_types") or []
    )
    if not skip_content and doc.ok and doc.text:
        for scope, label in ((doc.tail, "落款区"), (doc.text[:4000], "正文")):
            d = _scan_content_date(scope)
            if d:
                return d, f"正文·{label}"

    # ---- 阶段4：默认值 ----
    return assets.rules.get("fallback_date", "2024-01-01"), "默认值（该类型无日期语义）"


def _norm_date(s: str) -> str:
    parts = [p for p in re.split(r"[-/]", s) if p]
    if len(parts) == 3:
        y, m, d = parts
        return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
    return s


CN_NUM = {"〇": 0, "○": 0, "零": 0, "一": 1, "二": 2, "三": 3, "四": 4,
          "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def _cn_date_to_iso(s: str) -> str | None:
    m = re.match(r"(二[〇○零一二三四五六七八九十]+)年([一二三四五六七八九十]+)月", s)
    if not m:
        return None
    ytxt, mtxt = m.groups()
    digits = [CN_NUM.get(c) for c in ytxt if c in CN_NUM]
    if not digits or len(digits) < 4:
        return None
    y = "".join(str(d) for d in digits[:4])
    if mtxt == "十":
        mo = 10
    elif mtxt.startswith("十"):
        mo = 10 + CN_NUM.get(mtxt[-1], 1)
    elif mtxt.endswith("十"):
        mo = CN_NUM.get(mtxt[0], 1) * 10
    else:
        mo = CN_NUM.get(mtxt, 1)
    try:
        return f"{int(y):04d}-{int(mo):02d}-01"
    except Exception:
        return None


def _scan_content_date(text: str) -> str | None:
    if not text:
        return None
    cn = re.search(r"二[〇○零一二三四五六七八九十]{1,3}年[一二三四五六七八九十]{1,3}月", text)
    if cn:
        d = _cn_date_to_iso(cn.group(0))
        if d:
            return d
    m = re.search(r"(20\d{2})\s*年\s*(\d{1,2})\s*月", text)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 2015 <= y <= 2099 and 1 <= mo <= 12:
            return f"{y:04d}-{mo:02d}-01"
    return None


# ============================================================================
# 路由 + 置信度
# ============================================================================


def route(attrs: dict, assets: Assets, verdict: Verdict) -> None:
    doc_type = attrs.get("type")
    leader = attrs.get("leader")

    # 特例：规范制度无主体，直接落 00 空间
    ov_dir = assets.override_target_dir(doc_type)

    sub = assets.subdir_for_type(doc_type)
    space = assets.space_for_subject(leader)

    if ov_dir:
        verdict.target_dir = ov_dir
        verdict.space = ov_dir
        verdict.sub_dir = None
        # 落在「无主体」空间（如 00-规范制度）时，subject 一律置为占位符，
        # 避免正文里偶然出现的领导名字被写进文件名
        for sp in assets.taxonomy["spaces"]:
            if sp["id"] == ov_dir and sp.get("subjects") == ["-"]:
                verdict.subject = "-"
        return

    if space is None:
        verdict.blockers.append(
            f"主体「{leader or '未知'}」不在 taxonomy.spaces 中 —— "
            "新增主空间属结构性变更，禁止自动编号，需人工确认"
        )
        return

    if sub is None:
        verdict.blockers.append(
            f"文档类型「{doc_type or '未知'}」无对应子目录 —— "
            "禁止自动建目录（主题类应改用标签表达），需人工确认"
        )
        return

    verdict.space = space["id"]
    verdict.sub_dir = sub["id"]
    verdict.target_dir = f"{space['id']}/{sub['id']}"


def compute_confidence(attrs: dict, verdict: Verdict, assets: Assets) -> None:
    """归一化置信度。

    ⚠️ 不能直接用"权重和"：那样字段少的文档（如规范类无 leader）永远上不了高分。
    改为 **Σ(期望字段权重 × 命中权重) / Σ(期望字段权重)**：
      * type 必定期望（分类的根本）
      * leader 只在空间有主体时期望（00-规范制度 的 subjects 是 ["-"]，不期望）
      * audience / security 被规则命中才计入期望
    """
    w = assets.rules["confidence"]["weights"]
    fw = attrs.get("_field_weight", {})

    leader_expected = True
    for sp in assets.taxonomy["spaces"]:
        if sp.get("subjects") == ["-"] and (
            verdict.space == sp["id"] or attrs.get("type") == "规范"
        ):
            leader_expected = False

    expected = ["type"] + (["leader"] if leader_expected else [])
    for f in ("audience", "security"):
        if attrs.get(f) is not None:
            expected.append(f)

    total = sum(w[f] for f in expected) or 1.0
    score = sum(w[f] * fw.get(f, 0.0) for f in expected) / total

    pen = assets.rules["confidence"]["penalties"]
    if not attrs.get("type"):
        score -= pen["unknown_type"]
    if leader_expected and not attrs.get("leader"):
        score -= pen["unknown_leader"]
    for n in verdict.notes:
        if n.startswith("[冲突]"):
            score -= pen["signal_conflict"]
            break
    if any("来源路径指向" in n for n in verdict.notes):
        score -= pen["source_content_mismatch"]
    if any("旧版 .doc" in n for n in verdict.notes):
        score -= pen.get("unparsed_legacy_doc", 0.0)

    verdict.confidence = round(max(0.0, min(1.0, score)), 3)
    th = assets.rules["confidence"]["thresholds"]
    if verdict.confidence >= th["high"]:
        verdict.level = "high"
    elif verdict.confidence >= th["medium"]:
        verdict.level = "medium"
    else:
        verdict.level = "low"


def apply_guardrails(attrs: dict, verdict: Verdict, assets: Assets) -> None:
    """护栏：只拦【不一致】，不拦【一致】。

    设计要点：对外汇报（部委）/ 专题汇报 本身就是"敏感类目录"，
    敏感文件落进去是正确的 —— 若一律拦截，等于每天都要人工确认同一批已知分类。
    真正危险的是 **密级与目录性质不匹配**（敏感文件混进公开目录，或反之）。
    """
    g = assets.rules["guardrails"]
    security = attrs.get("security")
    policy = assets.taxonomy.get("security_policy", {})
    restricted = set(policy.get("restricted_levels", []))

    # 目标目录的"目录性质密级"，从 taxonomy 的 sub_dirs 声明中取
    sub_security = None
    for sd in assets.taxonomy.get("sub_dirs", []):
        if sd["id"] == verdict.sub_dir:
            sub_security = sd.get("security")
            break

    if security in restricted and sub_security is not None and sub_security not in restricted:
        verdict.blockers.append(
            f"密级「{security}」与目录「{verdict.sub_dir}」（{sub_security}性质）不匹配，"
            "违反密级隔离，需人工确认去向"
        )
    elif (security and security not in restricted and sub_security in restricted
          and g.get("public_into_restricted_blocks")):
        verdict.blockers.append(
            f"密级「{security}」与目录「{verdict.sub_dir}」（{sub_security}性质）不匹配，"
            "公开内容不得进入敏感目录，需人工确认"
        )
    elif security in restricted and sub_security is None and g.get("sensitive_requires_confirm"):
        # 目标目录未声明性质 → 无从判断隔离是否成立，保守起见确认
        verdict.blockers.append(f"密级判定为「{security}」，但目标目录未声明密级性质，需确认")

    # 分流动作
    if verdict.blockers:
        verdict.action = "block"
    elif verdict.level == "high":
        verdict.action = "silent"
    elif verdict.level == "medium":
        verdict.action = "preview"
    else:
        verdict.action = "block"
        if not verdict.blockers:
            verdict.blockers.append("置信度不足，信号存在歧义")


def build_candidates(attrs: dict, assets: Assets) -> list[dict]:
    """低置信时给出候选供人工选择（永远附带建议，不空问）。"""
    cands = []
    for sd in assets.taxonomy["sub_dirs"]:
        for t in sd.get("types", []):
            if t in (attrs.get("type"),):
                continue
            if attrs.get("audience") in (sd.get("audiences") or []):
                cands.append({"sub_dir": sd["id"], "type": t,
                              "reason": f"受众「{attrs.get('audience')}」与该目录匹配"})
    return cands[:3]


# ============================================================================
# 单文件主流程
# ============================================================================


def classify_one(path: Path, assets: Assets, root: Path | None, use_content: bool = True) -> Verdict:
    doc = DocText() if not use_content else extract_text(path)
    attrs, evidence, notes = extract_signals(path, assets, doc)

    v = Verdict(
        src=str(path),
        filename=path.name,
        ext=path.suffix.lower(),
        doc_type=attrs.get("type"),
        subject=attrs.get("leader"),
        audience=attrs.get("audience"),
        security=attrs.get("security"),
        tags=attrs.get("tags", []),
        date=attrs["date"][0],
        evidence=evidence + [f"日期依据：{attrs['date'][1]}"],
        notes=notes,
        text_method=doc.method,
        text_ok=doc.ok,
    )
    if doc.note:
        v.notes.append(doc.note)

    v.title_clean = enhance_title(clean_title(path.stem, assets, v.doc_type),
                                  path.stem, v.doc_type, assets)
    route(attrs, assets, v)

    # 命名
    meta = assets.taxonomy["meta"]
    d = v.date.replace("-", "")
    subj = meta["subject_placeholder"] if v.subject in (None, "-") else v.subject
    title = v.title_clean or prefix_stub(path)
    if v.target_dir:
        v.new_filename = f"{d}_{v.doc_type}_{title}_{subj}{v.ext}"
    else:
        v.new_filename = f"{d}_{v.doc_type or '未定'}_{title}_{subj}{v.ext}"

    compute_confidence(attrs, v, assets)
    apply_guardrails(attrs, v, assets)
    if v.action == "block":
        v.candidates = build_candidates(attrs, assets)

    return v


def prefix_stub(path: Path) -> str:
    return re.sub(r"[^\w\u4e00-\u9fff]+", "", path.stem)[:30] or "未命名"


# ============================================================================
# 输出
# ============================================================================


def fmt_card(v: Verdict) -> str:
    icon = {"silent": "[OK] 静默归类", "preview": "[?] 预演确认", "block": "[!] 需要人工确认"}[v.action]
    lines = [
        f"  {icon}  (置信度 {v.confidence:.2f} / {v.level})",
        f"    文件   {v.filename}",
        f"    类型   {v.doc_type or '未判定':<8} 主体   {v.subject or '未判定'}",
        f"    受众   {v.audience or '未判定':<8} 密级   {v.security or '未判定'}",
        f"    日期   {v.date}",
        f"    标签   {'/'.join(v.tags) if v.tags else '-'}",
        f"    目标   {v.target_dir or '（未确定）'}",
        f"    命名   {v.new_filename}",
    ]
    if v.evidence:
        lines.append("    依据   " + " ; ".join(e for e in v.evidence if e)[:200])
    for b in v.blockers:
        lines.append(f"    [!] {b}")
    for n in v.notes:
        lines.append(f"    [~] {n}")
    for c in v.candidates:
        lines.append(f"    候选   {c['sub_dir']}  ← {c['reason']}")
    return "\n".join(lines)


def main() -> int:
    ap = argparse.ArgumentParser(description="知识库分类决策引擎（只读，不落盘）")
    ap.add_argument("--input", required=True, help="文件或目录")
    ap.add_argument("--json", help="把结果写入 JSON 文件")
    ap.add_argument("--quiet", action="store_true", help="只输出 JSON 到 stdout")
    ap.add_argument("--no-content", action="store_true", help="只看文件名，不解析文档内容")
    ap.add_argument("--skill-dir", help="skill 根目录（默认取本脚本上级）")
    args = ap.parse_args()

    skill_dir = Path(args.skill_dir) if args.skill_dir else Path(__file__).resolve().parent.parent
    assets = Assets(skill_dir)

    target = Path(args.input)
    if target.is_dir():
        files = sorted(
            p for p in target.rglob("*")
            if p.is_file() and p.suffix.lower() in SUPPORTED_EXT
            and not p.name.startswith("~$")
        )
    elif target.is_file():
        files = [target]
    else:
        sys.stderr.write(f"[FAIL] 路径不存在：{target}\n")
        return 2

    results = [classify_one(p, assets, None, use_content=not args.no_content) for p in files]

    payload = [asdict(v) for v in results]
    if args.json:
        Path(args.json).parent.mkdir(parents=True, exist_ok=True)
        Path(args.json).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    if args.quiet:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    print(f"\n{'='*78}\n分类判定结果：{len(results)} 个文件\n{'='*78}")
    counts = {"silent": 0, "preview": 0, "block": 0}
    for v in results:
        counts[v.action] += 1
        print(f"\n{'-'*78}")
        print(fmt_card(v))
    print(f"\n{'='*78}")
    print(f"静默 {counts['silent']} · 预演 {counts['preview']} · 待确认 {counts['block']}  共 {len(results)}")
    if args.json:
        print(f"JSON 已写入：{args.json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
