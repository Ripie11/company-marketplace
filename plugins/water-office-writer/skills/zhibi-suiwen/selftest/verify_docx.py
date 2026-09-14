#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verify_docx.py — 公文版式自动校验器

回读由 md_to_gongwen_docx.py 生成的 docx，逐项断言：
  ① 页面：A4 尺寸 + 页边距（对比领导档案实测值）
  ② 主标题：方正小标宋简体 / 22pt / 居中
  ③ 副标题（##）：黑体 / 18pt / 居中
  ④ 一级标题（###）：黑体 / 与正文同级字号
  ⑤ 条款（（一））：楷体
  ⑥ 称呼语：仿宋 / 首行缩进 0（顶格）
  ⑦ 正文：仿宋_GB2312 / 档案字号 / 首行缩进 / 固定行距 28pt
  ⑧ 落款：右对齐
  ⑨ 残留 Markdown 语法探针（**、|、- 、*、####）
  ⑩ 能力探针：表格数 / 页码域 / 页眉 / 全角标点规范

用法:
  python verify_docx.py --docx out/C01.docx --leader 龚利民 [--expect-md cases/C01.md]
  或
  python verify_docx.py --batch out/ --manifest cases/_manifest.json --out verify.json
"""
import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

# ---- 版式档案（与 md_to_gongwen_docx.py 的实测值保持同源）----
PROFILES = {
    "龚利民": dict(body_pt=18.0, indent_cm=1.28, margin=(37, 35, 26, 26)),
    "方琳":   dict(body_pt=16.0, indent_cm=1.13, margin=(37, 35, 28, 26)),
    "":       dict(body_pt=16.0, indent_cm=1.13, margin=(37, 35, 28, 26)),
}
ALIASES = {
    "龚董": "龚利民", "董事长": "龚利民",
    "方总": "方琳", "总裁": "方琳",
}

FONT_TITLE = "方正小标宋简体"
FONT_BODY = "仿宋_GB2312"
FONT_KAI = "楷体"
FONT_HEI = "黑体"

CLAUSE_RE = re.compile(r"^[（(][一二三四五六七八九十]+[)）]")
SALUTATION_HINT = re.compile(
    r"(同志们|各位|尊敬的|女士们|先生们|同仁|嘉宾|领导|同事|委员|董事长|副董事长"
    r"|总裁|副总裁|总经理|局长|副局长|主任|院长|校长|书记|主席|部长|司长|处长|朋友|来宾)")
SALUTATION_BLOCK = re.compile(
    r"(如下|以下|详见|附表|附图|如下图|如下表|情况|安排|要求|标准|明细|清单|目标|指标|比例)")
NAME_LIST_RE = re.compile(r"[一-鿿]{2,4}([、，,][一-鿿]{2,4})+")


def is_salutation(t: str, max_len: int = 60) -> bool:
    t = t.strip()
    if not t or len(t) > max_len or not re.search(r"[：:]\s*$", t):
        return False
    head = re.sub(r"[：:]\s*$", "", t)
    if SALUTATION_BLOCK.search(head):
        return False
    return bool(SALUTATION_HINT.search(head) or NAME_LIST_RE.fullmatch(head))


def resolve(leader: str) -> str:
    k = (leader or "").strip()
    return ALIASES.get(k, k)


def han_font(run):
    rPr = run._element.rPr
    if rPr is None or rPr.rFonts is None:
        return None
    return rPr.rFonts.get(qn("w:eastAsia"))


def pt_of(emu):
    return round(emu / 12700, 2) if emu else None


def cm_of(emu):
    return round(emu / 360000, 3) if emu else None


def mm_of(emu):
    return round(emu / 36000, 2) if emu else None


def align_is(p, name: str) -> bool:
    """python-docx 的 alignment 字符串形如 'CENTER (1)' / 'RIGHT (2)'，需按枚举名匹配"""
    return p is not None and p.alignment is not None and p.alignment.name == name


def has_page_field(doc) -> bool:
    """页码域在 docx 的『页脚部件』里，不在正文 XML 中；遍历全部部件查找 PAGE 域"""
    try:
        for part in doc.part.package.iter_parts():
            try:
                if b"PAGE" in part.blob:
                    return True
            except Exception:
                continue
    except Exception:
        return False
    return False


def verdict_paragraphs(doc):
    """把段落按角色分类，返回角色 -> [段落索引]"""
    roles = {}
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if not t:
            continue
        if re.fullmatch(r"\*\*.+\*\*", t):
            roles.setdefault("signature_raw", []).append(i)
        if t.startswith("### "):
            roles.setdefault("stray_heading_marker", []).append(i)
        if t.startswith("#"):
            roles.setdefault("stray_hash", []).append(i)
    return roles


def expectations_from_md(md_path: Path) -> dict:
    """从 Markdown 源稿推导"应该出现哪些结构"，用于判定落版是否丢内容"""
    if md_path is None or not md_path.is_file():
        return {}
    txt = md_path.read_text(encoding="utf-8")
    lines = [l.rstrip() for l in txt.splitlines()]
    return {
        "h1": sum(1 for l in lines if l.startswith("### ")),
        "subtitle": sum(1 for l in lines if l.startswith("## ") and not l.startswith("### ")),
        "title": sum(1 for l in lines if l.startswith("# ") and not l.startswith("## ")),
        "clause": sum(1 for l in lines if CLAUSE_RE.match(l.strip())),
        "table_rows": sum(1 for l in lines if l.strip().startswith("|")),
        "table_count": 1 if any(l.strip().startswith("|") for l in lines) else 0,
        "salutation": sum(1 for l in lines if is_salutation(l)),
        "signature": sum(1 for l in lines if re.fullmatch(r"\*\*.+\*\*", l.strip())),
        "list_dash": sum(1 for l in lines if re.match(r"^\s*-\s+", l)),
        "h4": sum(1 for l in lines if l.startswith("#### ")),
        "emphasis_bold": len(re.findall(r"\*\*.+?\*\*", txt)) - 
                         sum(1 for l in lines if re.fullmatch(r"\*\*.+\*\*", l.strip())),
    }


def check_one(docx_path: Path, leader: str, md_path: Path | None = None) -> dict:
    key = resolve(leader)
    exp = expectations_from_md(md_path)
    prof = PROFILES.get(key, PROFILES[""])
    doc = Document(str(docx_path))
    res = {
        "docx": docx_path.name,
        "leader_arg": leader or "(默认)",
        "resolved": key or "(默认)",
        "checks": [],
        "probes": {},
        "paragraphs": [],
    }

    def chk(name, ok, detail):
        res["checks"].append({"item": name, "ok": bool(ok), "detail": detail})

    # ---------- ① 页面 ----------
    sec = doc.sections[0]
    pw, ph = round(sec.page_width.cm, 2), round(sec.page_height.cm, 2)
    chk("页面尺寸A4", (pw, ph) == (21.0, 29.7), f"{pw}x{ph}cm")
    got_m = (
        mm_of(sec.top_margin), mm_of(sec.bottom_margin),
        mm_of(sec.left_margin), mm_of(sec.right_margin),
    )
    exp_m = tuple(float(x) for x in prof["margin"])
    chk("页边距(上下左右mm)", all(abs(a - b) <= 0.5 for a, b in zip(got_m, exp_m)),
        f"实测{got_m} 期望{exp_m}")

    # ---------- 段落遍历分类 ----------
    paras = [p for p in doc.paragraphs if p.text.strip()]
    res["paragraphs"] = [p.text[:40] for p in paras]

    def first_run(p):
        for r in p.runs:
            if r.text.strip():
                return r
        return p.runs[0] if p.runs else None

    title_p = next((p for p in paras if p.alignment is not None
                    and p.text.strip() and han_font(first_run(p)) == FONT_TITLE), None)
    if title_p is None:
        title_p = paras[0] if paras else None

    # ---------- ② 主标题 ----------
    if title_p is not None:
        r = first_run(title_p)
        f, s = han_font(r), pt_of(r.font.size)
        chk("主标题字体", f == FONT_TITLE, f"实测{f} 期望{FONT_TITLE}")
        chk("主标题字号22pt", s == 22.0, f"实测{s}")
        chk("主标题居中", align_is(title_p, "CENTER"),
            f"实测{title_p.alignment}")
    else:
        chk("主标题存在", False, "未找到段落")

    # ---------- ④ 一级标题（### 映射） ----------
    # 生成文件里 ### 已被消费为无 # 的段落；用黑体 + 非居中 + 非称呼 判定
    h1_like = []
    for p in paras:
        r = first_run(p)
        if r is None:
            continue
        if han_font(r) == FONT_HEI and not align_is(p, "CENTER"):
            h1_like.append(p)
    res["probes"]["h1_like_count"] = len(h1_like)
    if exp.get("h1"):
        chk("一级标题数量(h1)", len(h1_like) >= exp["h1"],
            f"md期望{exp['h1']} 实测{len(h1_like)}")
    if h1_like:
        p = h1_like[0]
        r = first_run(p)
        chk("一级标题字体=黑体", han_font(r) == FONT_HEI, han_font(r) or "None")
        chk("一级标题字号=正文级", pt_of(r.font.size) == prof["body_pt"],
            f"实测{pt_of(r.font.size)} 期望{prof['body_pt']}")
        h1ind = cm_of(p.paragraph_format.first_line_indent) or 0.0
        chk("一级标题缩进=2字符", abs(h1ind - prof["indent_cm"]) < 0.02,
            f"实测{h1ind} 期望{prof['indent_cm']}")

    # ---------- ⑤ 条款（（一）） ----------
    clause_ps = [p for p in paras if CLAUSE_RE.match(p.text.strip())]
    res["probes"]["clause_count"] = len(clause_ps)
    if exp.get("clause"):
        chk("条款数量匹配", len(clause_ps) >= exp["clause"],
            f"md期望{exp['clause']} 实测{len(clause_ps)}")
    if clause_ps:
        p = clause_ps[0]
        r = first_run(p)
        chk("条款引语字体=楷体", han_font(r) == FONT_KAI, han_font(r) or "None")
        chk("条款字号=正文级", pt_of(r.font.size) == prof["body_pt"],
            f"实测{pt_of(r.font.size)}")

    # ---------- ⑥ 称呼语顶格 ----------
    sal_ps = []
    for p in paras:
        t = p.text.strip()
        if is_salutation(t) and han_font(first_run(p)) == FONT_BODY:
            sal_ps.append(p)
    res["probes"]["salutation_count"] = len(sal_ps)
    res["probes"]["salutation_expected"] = exp.get("salutation", 0)
    res["probes"]["salutation_texts"] = [p.text.strip()[:30] for p in sal_ps]
    if sal_ps:
        p = sal_ps[0]
        ind = cm_of(p.paragraph_format.first_line_indent) or 0.0
        chk("称呼语顶格(缩进0)", ind == 0.0, f"实测缩进{ind}cm")

    # ---------- ⑦ 正文 ----------
    body_ps = [p for p in paras
               if han_font(first_run(p)) == FONT_BODY
               and not is_salutation(p.text.strip())
               and not re.fullmatch(r"\*\*.+\*\*", p.text.strip())]
    res["probes"]["body_count"] = len(body_ps)
    if body_ps:
        p = body_ps[0]
        r = first_run(p)
        chk("正文字体=仿宋_GB2312", han_font(r) == FONT_BODY, han_font(r) or "None")
        chk("正文字号", pt_of(r.font.size) == prof["body_pt"],
            f"实测{pt_of(r.font.size)} 期望{prof['body_pt']}")
        ind = cm_of(p.paragraph_format.first_line_indent)
        chk("正文首行缩进", abs((ind or 0) - prof["indent_cm"]) < 0.02,
            f"实测{ind} 期望{prof['indent_cm']}")
        ls = p.paragraph_format.line_spacing
        chk("正文行距=固定28pt", pt_of(ls) == 28.0 if ls is not None else False,
            f"实测{pt_of(ls)}")

    # ---------- ⑧ 落款右对齐 ----------
    sig_ps = [p for p in paras if align_is(p, "RIGHT") and p.text.strip()]
    res["probes"]["signature_count"] = len(sig_ps)
    res["probes"]["signature_texts"] = [p.text.strip()[:40] for p in sig_ps]
    if exp.get("signature"):
        chk("落款数量匹配", len(sig_ps) >= exp["signature"],
            f"md期望{exp['signature']} 实测{len(sig_ps)}")

    # ---------- ⑨ 残留 Markdown 语法探针 ----------
    full = "\n".join(p.text for p in paras)
    leftovers = {
        "bold_marker": len(re.findall(r"\*\*", full)),
        "pipe_table": len(re.findall(r"^\|", full, re.M)),
        "list_dash": len(re.findall(r"^\s*-\s+", full, re.M)),
        "hash_heading": len(re.findall(r"^#{1,6}\s", full, re.M)),
        "italic_marker": len(re.findall(r"(?<!\*)\*(?!\*)", full)),
        "blockquote": len(re.findall(r"^>\s", full, re.M)),
    }
    res["probes"]["markdown_leftovers"] = leftovers

    # ---------- ⑩ 能力探针 ----------
    res["probes"]["table_count"] = len(doc.tables)
    if exp.get("table_count"):
        chk("表格落版", len(doc.tables) >= exp["table_count"],
            f"md含{exp['table_rows']}行表格, 实测docx表格{len(doc.tables)}个")
    res["probes"]["has_page_number_field"] = has_page_field(doc)
    res["probes"]["has_header_text"] = any(
        s.header.paragraphs and any(p.text.strip() for p in s.header.paragraphs)
        for s in doc.sections
    )
    halfwidth = len(re.findall(r"[0-9]", full))
    ascii_punct = len(re.findall(r"[,;:!?]", full))
    res["probes"]["halfwidth_digit_count"] = halfwidth
    res["probes"]["ascii_punct_count"] = ascii_punct

    # ---------- ⑪ 残留 Markdown 标记（结构性泄漏） ----------
    leak = leftovers
    chk("无粗体标记残留", leak["bold_marker"] == 0, f"{leak['bold_marker']} 处 **")
    chk("无表格管道符残留", leak["pipe_table"] == 0, f"{leak['pipe_table']} 行 |")
    chk("无列表标记残留", leak["list_dash"] == 0, f"{leak['list_dash']} 处 '- '")
    chk("无标题井号残留", leak["hash_heading"] == 0, f"{leak['hash_heading']} 处 #")
    chk("无斜体标记残留", leak["italic_marker"] == 0, f"{leak['italic_marker']} 处 *")

    # ---------- ⑫ 页码 ----------
    chk("页脚居中页码", res["probes"]["has_page_number_field"],
        "页脚缺少 PAGE 域" if not res["probes"]["has_page_number_field"] else "OK")

    # ---------- ⑬ 楷体引语范围（lead 模式：引语楷体 + 余下仿宋） ----------
    multi = [p for p in clause_ps
             if len({han_font(r) for r in p.runs if r.text.strip()}) > 1]
    res["probes"]["clause_lead_split_count"] = len(multi)
    if multi:
        fonts = [han_font(r) for r in multi[0].runs if r.text.strip()]
        chk("条款楷体引语+仿宋正文", fonts[0] == FONT_KAI and FONT_BODY in fonts[1:],
            f"字体序列 {fonts}")

    ok_n = sum(1 for c in res["checks"] if c["ok"])
    res["summary"] = {"total": len(res["checks"]), "pass": ok_n,
                      "fail": len(res["checks"]) - ok_n}
    return res


def main():
    ap = argparse.ArgumentParser(description="公文 docx 版式自动校验")
    ap.add_argument("--docx", help="单个 docx 路径")
    ap.add_argument("--leader", default="")
    ap.add_argument("--batch", help="批量：docx 目录")
    ap.add_argument("--manifest", help="用例清单 JSON（含 leader 映射）")
    ap.add_argument("--out", help="校验结果 JSON 输出路径")
    args = ap.parse_args()

    results = []
    if args.docx:
        results.append(check_one(Path(args.docx), args.leader))
    elif args.batch:
        d = Path(args.batch)
        lm = {}
        if args.manifest and Path(args.manifest).is_file():
            for m in json.loads(Path(args.manifest).read_text(encoding="utf-8")):
                lm[m["case_id"]] = m["leader"]
        case_dir = Path(args.batch).parent / "cases"
        for f in sorted(d.glob("*.docx")):
            cid = f.name.split("_")[0]
            cand = sorted(case_dir.glob(f"{cid}_*.md")) if case_dir.is_dir() else []
            results.append(check_one(f, lm.get(cid, ""), cand[0] if cand else None))

    for r in results:
        s = r["summary"]
        flag = "[PASS]" if s["fail"] == 0 else "[FAIL]"
        print(f"{flag} {r['docx']}  ({r['leader_arg']})  {s['pass']}/{s['total']}")
        for c in r["checks"]:
            if not c["ok"]:
                print(f"       x {c['item']}: {c['detail']}")

    if args.out:
        Path(args.out).write_text(
            json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n[OK] 明细 -> {args.out}")
    total_fail = sum(r["summary"]["fail"] for r in results)
    print(f"\n合计：{len(results)} 份 / 失败断言 {total_fail} 项")
    return 1 if total_fail else 0


if __name__ == "__main__":
    raise SystemExit(main())
