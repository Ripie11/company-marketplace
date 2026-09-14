#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""backtest.py —— 分类引擎回测台

【为什么必须有这个脚本】
分类规则（routing-rules.yaml）是可以被不断微调的。没有回测，"调优"就是凭感觉；
有了回测，每次改规则都能立刻看到命中率是涨还是跌。**规则改动必须跑这个脚本。**

流程：
    人工真值（_metadata.csv，GBK）  ┐
                                    ├─→ classify_one() 逐项比对 → 命中率报告
    待判文件（原始素材目录）        ┘

用法：
    python backtest.py --truth <_metadata.csv> --input <原始素材目录> [--json 报告.json]
    python backtest.py --truth 知识库/_metadata.csv --input "_backtest/知识库 - 原文件"

退出码：0 达标（type/leader/dir 均 ≥ 阈值）；1 未达标；2 参数错误。
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import sys
import difflib
import re
from collections import Counter
from pathlib import Path

if sys.stdout.encoding is None or sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass

sys.path.insert(0, str(Path(__file__).resolve().parent))
from classify import Assets, classify_one, SUPPORTED_EXT  # noqa: E402

# 磁盘上曾出现过的旧空间名 → 规范空间名（真值表与磁盘都按规范名比对）
SPACE_ALIASES = {
    "02-方琳（执行总裁）": "02-方琳（总裁）",
    "01-龚利民": "01-龚利民（董事长）",
}

# 达标线
THRESHOLDS = {"type": 0.95, "leader": 0.95, "dir": 0.90}


def norm_dir(d: str | None) -> str:
    """真值表的 new_path 是【完整路径含文件名】，取目录部分再规范化。

    同时吸收历史上的空间名变体（磁盘曾是「02-方琳（总裁）」）。
    """
    if not d:
        return ""
    parts = d.replace("\\", "/").rstrip("/").split("/")
    # 末段若形如 xxx.ext，说明是文件名，去掉
    if parts and re.search(r"\.[A-Za-z0-9]{1,5}$", parts[-1]):
        parts = parts[:-1]
    parts = [SPACE_ALIASES.get(p, p) for p in parts]
    return "/".join(parts)


def norm_date(d: str | None) -> str:
    """统一为 YYYY-MM-DD（真值表用的是未补零的 2026/7/23）。"""
    if not d:
        return ""
    m = re.match(r"\s*(20\d{2})[/\-.](\d{1,2})[/\-.](\d{1,2})", d)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return d.replace("/", "-").strip()


def title_sim(a: str, b: str) -> float:
    a = re.sub(r"[\s\-_（）()【】]", "", a or "")
    b = re.sub(r"[\s\-_（）()【】]", "", b or "")
    if not a and not b:
        return 1.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def load_truth(csv_path: Path) -> list[dict]:
    """真值表是 GBK 编码 —— 这不是笔误，是历史遗留，必须用 gb18030 读。"""
    for enc in ("gb18030", "utf-8-sig", "utf-8"):
        try:
            with open(csv_path, encoding=enc) as f:
                rows = list(csv.DictReader(f))
            if rows and "new_path" in rows[0]:
                return rows
        except (UnicodeDecodeError, KeyError):
            continue
    raise SystemExit(f"[FAIL] 无法解析真值表（编码不是 gb18030/utf-8）：{csv_path}")


def main() -> int:
    ap = argparse.ArgumentParser(description="分类引擎回测台")
    ap.add_argument("--truth", required=True, help="人工真值 _metadata.csv")
    ap.add_argument("--input", required=True, help="原始素材目录（用于重新分类）")
    ap.add_argument("--json", help="把逐项比对结果写入 JSON")
    ap.add_argument("--skill-dir", help="skill 根目录")
    ap.add_argument("--no-content", action="store_true", help="只按文件名分类（快速回归）")
    args = ap.parse_args()

    skill_dir = Path(args.skill_dir).resolve() if args.skill_dir else Path(__file__).resolve().parent.parent
    assets = Assets(skill_dir)
    truth = load_truth(Path(args.truth))

    # 建索引：真值表的 src_rel basename → 真值行
    by_name: dict[str, dict] = {}
    for r in truth:
        base = Path(r["src_rel"].replace("\\", "/")).name
        by_name[base] = r

    root = Path(args.input)
    files = sorted(
        p for p in root.rglob("*")
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXT and not p.name.startswith("~$")
    )

    rows: list[dict] = []
    unmatched: list[str] = []

    for p in files:
        t = by_name.get(p.name)
        if t is None:
            unmatched.append(p.name)
            continue
        v = classify_one(p, assets, None, use_content=not args.no_content)

        got_dir = norm_dir(v.target_dir)
        exp_dir = norm_dir(t["new_path"])
        got_date, exp_date = norm_date(v.date), norm_date(t["date"])
        sim = title_sim(v.title_clean, t["title_clean"])

        rows.append({
            "file": p.name,
            "hit_type": v.doc_type == t["type"],
            "hit_leader": (v.subject or "-") == (t["leader"] or "-"),
            "hit_dir": got_dir == exp_dir,
            "hit_audience": (v.audience or "") == (t["audience"] or ""),
            "hit_security": (v.security or "") == (t["security"] or ""),
            "date_exact": got_date == exp_date,
            "date_month": got_date[:7] == exp_date[:7],
            "date_year": got_date[:4] == exp_date[:4],
            "exp_type": t["type"], "got_type": v.doc_type,
            "exp_leader": t["leader"], "got_leader": v.subject,
            "exp_dir": exp_dir, "got_dir": got_dir,
            "exp_date": exp_date, "got_date": got_date,
            "exp_title": t["title_clean"], "got_title": v.title_clean,
            "title_sim": round(sim, 3),
            "exp_name": t["new_filename"], "got_name": v.new_filename,
            "action": v.action, "confidence": v.confidence,
        })

    # ---------------- 汇总 ----------------
    n = len(rows)
    if n == 0:
        sys.stderr.write("[FAIL] 没有任何文件与真值表匹配，检查 --input 是否指向原始素材目录\n")
        return 2

    def rate(k: str) -> float:
        return sum(1 for r in rows if r[k]) / n

    type_acc = rate("hit_type")
    leader_acc = rate("hit_leader")
    dir_acc = rate("hit_dir")
    aud_acc = rate("hit_audience")
    sec_acc = rate("hit_security")
    d_exact = rate("date_exact")
    d_month = rate("date_month")
    d_year = rate("date_year")
    t_sim = sum(r["title_sim"] for r in rows) / n
    t_ok = sum(1 for r in rows if r["title_sim"] >= 0.6) / n
    actions = Counter(r["action"] for r in rows)

    print(f"\n{'='*82}")
    print(f"分类引擎回测报告  ·  {n} 个文件  ·  {'文件名-only' if args.no_content else '文件名+正文'}")
    print(f"{'='*82}")
    print(f"  {'字段':<14}{'命中':>7}{'准确率':>10}   {'说明'}")
    print(f"  {'-'*74}")
    def line(label, val, note, extra=""):
        flag = "OK " if val is not None else "-- "
        print(f"  {label:<14}{'' if val is None else f'{val:.1%}':>12}   {note}{extra}")

    print(f"  {'type':<14}{type_acc:>12.1%}   {'硬指标（阈值 %.0f%%）' % (THRESHOLDS['type']*100)}")
    print(f"  {'leader':<14}{leader_acc:>12.1%}   {'硬指标（阈值 %.0f%%）' % (THRESHOLDS['leader']*100)}")
    print(f"  {'target_dir':<14}{dir_acc:>12.1%}   {'硬指标（阈值 %.0f%%）' % (THRESHOLDS['dir']*100)}")
    print(f"  {'audience':<14}{aud_acc:>12.1%}")
    print(f"  {'security':<14}{sec_acc:>12.1%}")
    print(f"  {'date(精确)':<14}{d_exact:>12.1%}   同月 {d_month:.1%} · 同年 {d_year:.1%}")
    print(f"  {'title(相似)':<14}{t_sim:>12.3f}   相似度≥0.60 的占 {t_ok:.1%}")
    print(f"  {'-'*74}")
    print(f"  分流动作：静默 {actions.get('silent',0)} · 预演 {actions.get('preview',0)} · 待确认 {actions.get('block',0)}")

    # ---------------- 逐项差异 ----------------
    print(f"\n{'='*82}\n逐项差异\n{'='*82}")
    for r in rows:
        miss = [k for k in ("hit_type", "hit_leader", "hit_dir") if not r[k]]
        if not miss:
            continue
        tags = []
        if not r["hit_type"]:
            tags.append(f"type {r['got_type']} ≠ {r['exp_type']}")
        if not r["hit_leader"]:
            tags.append(f"leader {r['got_leader']} ≠ {r['exp_leader']}")
        if not r["hit_dir"]:
            tags.append(f"dir {r['got_dir']} ≠ {r['exp_dir']}")
        print(f"\n  [X] {r['file'][:70]}")
        print(f"      " + " ; ".join(tags))
        if not r["date_exact"]:
            print(f"      date {r['got_date']} ≠ {r['exp_date']}（同年={r['date_year']} 同月={r['date_month']}）")
        if r["title_sim"] < 0.75:
            print(f"      title 相似 {r['title_sim']:.2f}")
            print(f"        期望 {r['exp_title']}")
            print(f"        实得 {r['got_title']}")
            print(f"        真值命名 {r['exp_name']}")
            print(f"        实得命名 {r['got_name']}")

    if unmatched:
        print(f"\n  [i] {len(unmatched)} 个文件未在真值表中找到：{[u[:30] for u in unmatched[:5]]}")

    # 仅日期不一致（分类正确但日期有偏差）—— 单列，便于判断是否值得回修
    date_bad = [r for r in rows if not r["date_exact"]]
    if date_bad:
        print(f"\n{'='*82}\n日期偏差（{len(date_bad)} 项；分类均已命中，仅日期不同）\n{'='*82}")
        for r in date_bad:
            flag = "同年" if r["date_year"] else "跨年"
            print(f"  · {r['file'][:58]}")
            print(f"      实得 {r['got_date']}  ≠  真值 {r['exp_date']}   [{flag}]")

    # ---------------- 达标判定 ----------------
    passed = type_acc >= THRESHOLDS["type"] and leader_acc >= THRESHOLDS["leader"] and dir_acc >= THRESHOLDS["dir"]
    print(f"\n{'='*82}")
    print(f"结论：{'✅ 达标（可推进 L3 乐享同步）' if passed else '❌ 未达标 —— 需回修 routing-rules.yaml 后重跑'}")
    print(f"{'='*82}\n")

    if args.json:
        Path(args.json).parent.mkdir(parents=True, exist_ok=True)
        Path(args.json).write_text(json.dumps({
            "summary": {
                "n": n, "type": type_acc, "leader": leader_acc, "dir": dir_acc,
                "audience": aud_acc, "security": sec_acc,
                "date_exact": d_exact, "date_month": d_month, "date_year": d_year,
                "title_sim_mean": round(t_sim, 3), "title_sim_ok": t_ok,
                "actions": dict(actions), "passed": passed,
            },
            "rows": rows,
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON 已写入：{args.json}")

    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
