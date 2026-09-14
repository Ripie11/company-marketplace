#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_pdf.py - 从 .pdf 文件提取格式与文本（v3 加 OCR 通道）

设计要点：
- v1: pdfplumber 字符级 + y 坐标聚类（50 行雏形）
- v2 (2026-09-11): FONT_ALIASES 渲染字体归一 + SIZE_LEVEL 字号反推 + 页眉页脚过滤 + 段落重构（341 行）
- v3 (2026-09-11): 加扫描件检测 + RapidOCR 通道
    - 每页先 pdfplumber 取文本：长度 < 阈值 → 视为扫描件 → 渲染为图片 → RapidOCR(中英)
    - OCR 结果 bbox → 转 pdfplumber chars 格式（估算字号 = bbox 高度 × 0.75）→ 后续聚类逻辑零改动
    - OCR 通道 graceful 降级：rapidocr_onnxruntime 未装 → OCR_AVAILABLE=False，仅跑文本层（不影响主流程）
    - 每页加 text_source 标记 / 整体 result 加 meta.ocr_pages 计数

依赖:
    必需: pip install pdfplumber
    可选: pip install rapidocr-onnxruntime pillow  (扫描件 PDF 才需要)
"""
import argparse
import json
import os
import re
from collections import Counter
from pathlib import Path

import pdfplumber

# ---------- OCR 通道（可选依赖，graceful 降级） ----------
try:
    from rapidocr_onnxruntime import RapidOCR
    _OCR_ENGINE = RapidOCR()
    OCR_AVAILABLE = True
except Exception as _e:
    _OCR_ENGINE = None
    OCR_AVAILABLE = False
    _OCR_IMPORT_ERROR = str(_e)

from PIL import Image
import io
import numpy as np


# ---------- 字体归一表（CJK 公文常用字体 → 标准名） ----------
FONT_ALIASES = {
    # 宋体族
    "SimSun": "宋体", "宋体": "宋体", "宋体-简": "宋体", "宋体-简体": "宋体",
    "STSong": "宋体", "STSong-Light": "宋体", "NSimSun": "宋体",
    # 黑体族
    "SimHei": "黑体", "黑体": "黑体", "Heiti SC": "黑体", "STHeiti": "黑体",
    # 仿宋族
    "FangSong": "仿宋_GB2312", "FangSong_GB2312": "仿宋_GB2312", "仿宋_GB2312": "仿宋_GB2312",
    "STFangsong": "仿宋_GB2312", "仿宋": "仿宋_GB2312",
    # 楷体族
    "KaiTi": "楷体_GB2312", "KaiTi_GB2312": "楷体_GB2312", "楷体_GB2312": "楷体_GB2312",
    "STKaiti": "楷体_GB2312", "楷体": "楷体_GB2312",
    # 方正小标宋
    "FZ XiaoBiaoSong-B05S": "方正小标宋简体", "方正小标宋简体": "方正小标宋简体",
    "FZXiaoBiaoSong-B05S": "方正小标宋简体", "FZXBiaoSong-B04S": "方正小标宋简体",
    "FZXBSJW--GB1-0": "方正小标宋简体",
}

# CJK 公文常用字号（pt → 名称），用于标题反推
SIZE_LEVEL = [
    (28, "main_title_lg"),   # 一号
    (22, "main_title"),      # 二号 - 主标题常用
    (18, "subtitle"),        # 小二 - 副标题 / 一级标题
    (16, "body_or_h1"),      # 三号 - 正文或一级标题
    (15, "body_or_h2"),
    (14, "body_or_h3"),      # 四号
    (12, "small"),
]

# 页眉页脚识别模式
PAGE_NO_PATTERN = re.compile(r"^[\s—\-—\-]*[—\-\d\s]{1,8}[\s—\-—\-]*(页|Page|P\.|／／|—\d+—)*$", re.I)
HEADER_FOOTER_KW = ["页码", "第 页", "Page ", "/P", "—", "打印日期", "印发"]

# OCR 触发阈值（每页 extract_text 字符串长度小于此视为扫描件）
SCAN_TEXT_THRESHOLD = 50


def normalize_font(name):
    """字体名归一"""
    if not name:
        return None
    return FONT_ALIASES.get(name, FONT_ALIASES.get(name.strip(), name))


def classify_size(size_pt):
    """字号 → 层级标记"""
    if size_pt is None:
        return None
    for threshold, level in SIZE_LEVEL:
        if size_pt >= threshold:
            return level
    return "small"


def is_header_footer(text, top, page_height):
    """判断是否页眉页脚（基于 y 坐标 + 文本模式）"""
    t = text.strip()
    if not t:
        return False
    # 顶部 50pt / 底部 50pt 阈值
    if top < 50 or top > page_height - 50:
        return True
    # 文本模式
    if PAGE_NO_PATTERN.match(t):
        return True
    for kw in HEADER_FOOTER_KW:
        if kw in t and len(t) < 30:
            return True
    return False


def is_scanned_page(page, threshold=SCAN_TEXT_THRESHOLD):
    """检测单页是否为扫描件（无文本层）"""
    try:
        text = page.extract_text() or ""
        # 再看 chars（pdfplumber 的字符级输出，扫描件通常为空）
        has_chars = len(page.chars) > 5
        return len(text.strip()) < threshold and not has_chars
    except Exception:
        return True  # 解析失败也按扫描件处理


def render_page_to_image(page, dpi=200):
    """把 PDF 单页渲染为 PIL Image（白底，RGB）"""
    img = page.to_image(resolution=dpi)
    pil_img = img.original  # PIL.Image
    if pil_img.mode != "RGB":
        pil_img = pil_img.convert("RGB")
    return pil_img


def ocr_page_to_chars(page, dpi=200):
    """扫描页 → RapidOCR → pdfplumber chars 格式

    Returns: (chars: list[dict], avg_confidence: float)
    """
    if not OCR_AVAILABLE:
        return [], 0.0

    pil_img = render_page_to_image(page, dpi=dpi)
    img_np = np.array(pil_img)
    ph = page.height  # PDF 坐标系的高度（pt）
    pw = page.width

    # RapidOCR 返回 [[[box], text, conf], ...]，box=[[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
    result, _elapsed = _OCR_ENGINE(img_np)

    chars = []
    confidences = []
    if not result:
        return [], 0.0

    for item in result:
        try:
            box, text, conf = item[0], item[1], float(item[2])
        except (IndexError, ValueError, TypeError):
            continue
        if not text or not text.strip():
            continue
        # bbox 坐标：左上、右上、右下、左下
        xs = [pt[0] for pt in box]
        ys = [pt[1] for pt in box]
        x0, x1 = min(xs), max(xs)
        y0, y1 = min(ys), max(ys)
        # 图片像素 → PDF 坐标系（pt）：page.width_px / pw 比例
        # pdfplumber.to_image 内部按 dpi 渲染，1 inch = 72 pt，所以 1 px = 72/dpi pt
        scale = 72.0 / dpi
        px0 = x0 * scale
        px1 = x1 * scale
        py0 = y0 * scale
        py1 = y1 * scale
        # 字号粗估 = bbox 高度 × 0.75（经验系数，匹配中文字符实际尺寸）
        bbox_h_pt = py1 - py0
        size_pt = round(bbox_h_pt * 0.75, 1)
        for ch in text:
            chars.append({
                "text": ch,
                "x0": px0,
                "x1": px1,
                "top": py0,
                "bottom": py1,
                "size": size_pt,
                "fontname": "OCR_Recognized",
                "from_ocr": True,
            })
        confidences.append(conf)

    return chars, (sum(confidences) / len(confidences)) if confidences else 0.0


def cluster_chars_to_lines(chars, y_tolerance=3):
    """字符按 (top, x0) 聚类成行。返回 list[dict]，每元素: {text, size_pt, fontname, x0, top, bottom}"""
    if not chars:
        return []
    sorted_chars = sorted(chars, key=lambda c: (round(c["top"], 1), c["x0"]))
    lines = []
    current = []
    cur_top = None
    for ch in sorted_chars:
        top = round(ch["top"], 1)
        if cur_top is None or abs(top - cur_top) <= y_tolerance:
            current.append(ch)
            cur_top = top if cur_top is None else cur_top
        else:
            lines.append(current)
            current = [ch]
            cur_top = top
    if current:
        lines.append(current)
    # 合并行为字符串
    out = []
    for line_chars in lines:
        line_chars = sorted(line_chars, key=lambda c: c["x0"])
        text = "".join(c["text"] for c in line_chars).strip()
        if not text:
            continue
        sizes = [c.get("size", 0) for c in line_chars if c.get("size")]
        fonts = [c.get("fontname", "") for c in line_chars if c.get("fontname")]
        out.append({
            "text": text,
            "size_pt": round(sum(sizes) / len(sizes), 1) if sizes else None,
            "fontname": normalize_font(fonts[0]) if fonts else None,
            "x0": min(c["x0"] for c in line_chars),
            "top": round(min(c["top"] for c in line_chars), 1),
            "bottom": round(max(c["bottom"] for c in line_chars), 1),
        })
    return out


def cluster_lines_to_paragraphs(lines, page_width):
    """行 → 段：基于字号变化（标题单独成段）、y 间距（> 1.5 倍行高）、首行缩进启发式"""
    if not lines:
        return []

    # 计算基准行高（众数）
    line_heights = Counter()
    for i in range(len(lines) - 1):
        gap = lines[i + 1]["top"] - lines[i]["top"]
        if 5 < gap < 50:
            line_heights[round(gap, 1)] += 1
    base_line_height = line_heights.most_common(1)[0][0] if line_heights else 14

    paragraphs = []
    cur_lines = []
    cur_body_size = None
    for i, ln in enumerate(lines):
        # 字号突变且与 body 不同 → 视为标题，单独成段
        if cur_body_size is None:
            cur_body_size = ln["size_pt"]
        is_title_like = (
            ln["size_pt"] and cur_body_size and abs(ln["size_pt"] - cur_body_size) > 2
        )
        # 与上一行的 y 间距
        if cur_lines and not is_title_like:
            prev = cur_lines[-1]
            gap = ln["top"] - prev["bottom"]
            if gap > base_line_height * 1.6:
                paragraphs.append(_merge_lines_to_paragraph(cur_lines))
                cur_lines = []
        cur_lines.append(ln)
        if is_title_like and i < len(lines) - 1:
            paragraphs.append(_merge_lines_to_paragraph(cur_lines))
            cur_lines = []
            cur_body_size = ln["size_pt"]
    if cur_lines:
        paragraphs.append(_merge_lines_to_paragraph(cur_lines))
    return paragraphs


def _merge_lines_to_paragraph(lines):
    """同一段内的多行合并：检测首行缩进 + 拼接"""
    if len(lines) == 1:
        ln = lines[0]
        return {
            "text": ln["text"],
            "size_pt": ln["size_pt"],
            "fontname": ln["fontname"],
            "is_first_line_indent": False,  # 单行无法判断
        }
    # 多行：判断首行 x0 是否大于第二行 → 缩进
    indent = lines[0]["x0"] > lines[1]["x0"] + 5
    text = "".join(ln["text"] for ln in lines)
    return {
        "text": text,
        "size_pt": lines[0]["size_pt"],
        "fontname": lines[0]["fontname"],
        "is_first_line_indent": indent,
    }


def extract_one_pdf(file_path, enable_ocr=True, ocr_dpi=200):
    """处理单个 PDF 文件"""
    result = {
        "file": os.path.basename(file_path),
        "file_type": "pdf",
        "format_profile": {"fonts": {}, "paragraph": {}, "page": {}, "headings": []},
        "text_profile": {},
        "warnings": [],
        "meta": {
            "ocr_available": OCR_AVAILABLE,
            "ocr_enabled": enable_ocr and OCR_AVAILABLE,
            "ocr_pages": 0,
            "text_layer_pages": 0,
            "ocr_avg_confidence": None,
        },
    }

    try:
        with pdfplumber.open(str(file_path)) as pdf:
            page_count = len(pdf.pages)
            all_paragraphs = []
            all_lines = []  # 用于格式统计
            headings = []
            body_size_counter = Counter()
            body_font_counter = Counter()
            indent_counter = Counter()
            confidences_collected = []

            for page_idx, page in enumerate(pdf.pages):
                ph = page.height
                # ---------- 1. 选择数据源：文本层 or OCR ----------
                use_ocr = enable_ocr and OCR_AVAILABLE and is_scanned_page(page)
                if use_ocr:
                    chars, avg_conf = ocr_page_to_chars(page, dpi=ocr_dpi)
                    result["meta"]["ocr_pages"] += 1
                    if avg_conf:
                        confidences_collected.append(avg_conf)
                    text_source = "ocr"
                else:
                    chars = [c for c in page.chars]
                    text_source = "text_layer"
                    result["meta"]["text_layer_pages"] += 1

                # 过滤页眉页脚
                kept = [c for c in chars if not is_header_footer(c["text"], c["top"], ph)]
                lines = cluster_chars_to_lines(kept)
                paragraphs = cluster_lines_to_paragraphs(lines, page.width)

                # 收集正文统计（首段非首行 + 段长 > 30 字视为正文）
                body_paragraphs = [p for p in paragraphs if len(p["text"]) > 30]
                for p in body_paragraphs:
                    if p["size_pt"]:
                        body_size_counter[round(p["size_pt"])] += 1
                    if p["fontname"]:
                        body_font_counter[p["fontname"]] += 1
                    if p["is_first_line_indent"]:
                        indent_counter["indent"] += 1
                    else:
                        indent_counter["no_indent"] += 1

                # 标题识别：段长 < 80 + 字号 > body_size 众数 × 1.2
                body_mode_size = body_size_counter.most_common(1)[0][0] if body_size_counter else 14
                for p in paragraphs:
                    text = p["text"]
                    if p["size_pt"] and p["size_pt"] >= body_mode_size * 1.2 and len(text) < 80:
                        level = classify_size(p["size_pt"])
                        headings.append({
                            "level": level,
                            "text": text[:60],
                            "size_pt": p["size_pt"],
                            "fontname": p["fontname"],
                            "page": page_idx + 1,
                            "text_source": text_source,
                        })

                all_paragraphs.extend([{"text": p["text"], "text_source": text_source} for p in paragraphs])
                all_lines.extend(lines)

            if confidences_collected:
                result["meta"]["ocr_avg_confidence"] = round(sum(confidences_collected) / len(confidences_collected), 3)

            # 格式档案
            if body_font_counter:
                result["format_profile"]["fonts"]["body"] = {
                    "family": body_font_counter.most_common(1)[0][0],
                    "size_pt": body_size_counter.most_common(1)[0][0] if body_size_counter else None,
                    "color": "000000",
                    "text_source_mix": "ocr+text_layer" if result["meta"]["ocr_pages"] and result["meta"]["text_layer_pages"] else ("ocr" if result["meta"]["ocr_pages"] else "text_layer"),
                }
                result["format_profile"]["fonts_dist"] = body_font_counter.most_common(6)
                result["format_profile"]["sizes_dist"] = body_size_counter.most_common(6)

            # 段落格式
            if indent_counter:
                indent_pct = indent_counter.get("indent", 0) / sum(indent_counter.values())
                result["format_profile"]["paragraph"] = {
                    "first_line_indent_inferred": indent_pct > 0.5,
                    "indent_distribution": dict(indent_counter),
                    "line_spacing": None,  # PDF 行距无可信信号
                    "line_spacing_rule": None,
                    "space_before_pt": None,
                    "space_after_pt": None,
                }
                if indent_pct <= 0.5:
                    result["warnings"].append("PDF 首行缩进比例 < 50%，缩进判断不可靠")

            # 页面（PDF 无 margin 信号，仅记尺寸）
            if pdf.pages:
                p0 = pdf.pages[0]
                result["format_profile"]["page"] = {
                    "paper_width_mm": round(p0.width * 25.4 / 72, 1),
                    "paper_height_mm": round(p0.height * 25.4 / 72, 1),
                    "margin_top_mm": None,
                    "margin_bottom_mm": None,
                    "margin_left_mm": None,
                    "margin_right_mm": None,
                }
                result["warnings"].append("PDF 页边距无可靠信号，已记 None")

            result["format_profile"]["headings"] = headings[:30]  # 最多 30 个

            # 文本档案
            full_text = "\n".join(p["text"] for p in all_paragraphs)
            paragraphs_clean = [p["text"] for p in all_paragraphs if p["text"].strip()]
            result["text_profile"] = {
                "total_chars": len(full_text),
                "paragraph_count": len(paragraphs_clean),
                "paragraphs": paragraphs_clean,
                "first_sentences": [p["text"][:50] for p in all_paragraphs[:10] if p["text"]],
                "last_sentences": [p["text"][-50:] for p in all_paragraphs if p["text"]][-5:],
            }

            result["page_count"] = page_count
            result["status"] = "success"

            # 顶层 OCR 信息
            if result["meta"]["ocr_pages"] == page_count and page_count > 0:
                result["warnings"].append(f"全部 {page_count} 页均为扫描件，已走 OCR（avg conf={result['meta']['ocr_avg_confidence']}）")
            elif result["meta"]["ocr_pages"] > 0:
                result["warnings"].append(f"混合 PDF：{result['meta']['ocr_pages']}/{page_count} 页走 OCR（avg conf={result['meta']['ocr_avg_confidence']}），其余走文本层")
            elif not OCR_AVAILABLE and enable_ocr:
                result["warnings"].append("启用 OCR 但 rapidocr_onnxruntime 未安装，已按文本层处理。安装命令: pip install rapidocr-onnxruntime")
    except Exception as e:
        result["status"] = "failed"
        result["error"] = str(e)

    return result


def main():
    parser = argparse.ArgumentParser(description="从 PDF 提取格式与文本（字符级 + 字号反推 + 扫描件自动 OCR）")
    parser.add_argument("--input", required=True, help="输入文件或目录")
    parser.add_argument("--output", required=True, help="输出 JSON 文件")
    parser.add_argument("--ocr", dest="enable_ocr", action="store_true", default=True,
                        help="启用 OCR 处理扫描件 PDF（默认开启，需 rapidocr-onnxruntime）")
    parser.add_argument("--no-ocr", dest="enable_ocr", action="store_false",
                        help="禁用 OCR（仅走文本层）")
    parser.add_argument("--ocr-dpi", type=int, default=200, help="OCR 渲染 DPI（默认 200，越大越慢但越准）")
    args = parser.parse_args()

    input_path = Path(args.input)
    if input_path.is_dir():
        files = sorted(input_path.glob("**/*.pdf"))
    else:
        files = [input_path]

    if not files:
        print(f"⚠️ 未找到 PDF 文件：{args.input}")
        return

    results = [extract_one_pdf(f, enable_ocr=args.enable_ocr, ocr_dpi=args.ocr_dpi) for f in files]

    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    success_count = sum(1 for r in results if r["status"] == "success")
    ocr_total = sum(r.get("meta", {}).get("ocr_pages", 0) for r in results)
    print(f"✅ 处理完成：{success_count}/{len(results)} 成功，输出至 {args.output}")
    if ocr_total:
        print(f"🔍 共 {ocr_total} 页走 OCR 通道")
    warn_count = sum(len(r.get("warnings", [])) for r in results)
    if warn_count:
        print(f"⚠️ 共 {warn_count} 条警告（见各文件 warnings 字段）")


if __name__ == "__main__":
    main()