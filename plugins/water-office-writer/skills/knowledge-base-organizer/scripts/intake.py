#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""intake.py —— 对话式增量接入（L2）

把「一份新文档该不该入库、放到哪、叫什么名字」变成一次**可预演、可回滚、可审计**
的对话，而不是一次性的批处理脚本。

    classify.py  只做判定（只读）
    intake.py    做决定 + 落盘（写文件、写台账）
    kbsync.py    把落好的盘同步到乐享（L3）

三种分流动作（由 classify.py 给出，intake.py 只负责尊重它）：

    [OK] silent   置信度 >= 0.80   → 默认自动接受
    [?]  preview  置信度 0.55~0.80 → 默认【不落盘】，等你点头
    [!]  block    置信度 < 0.55 或撞护栏 → 一律不落盘，必须人工定夺

【零风险铁律】
  1.  **只复制，永不移动/删除**源文件 —— 原始素材是唯一的真相来源
  2.  默认 dry-run，必须显式 `--apply` 才动盘（对应 routing-rules: dry_run_default）
  3.  目标已存在同内容 → 幂等复用；同路径不同内容 → 报冲突，绝不自动覆盖
  4.  每次落盘都写 `_intake_log.csv`，可回溯

用法：
    # 1) 先出计划（默认 dry-run，不写任何文件）
    python intake.py --input <新文件或目录> --kb-root <知识库根>

    # 2) 把计划落成 JSON，交给人/上游 Agent 复核
    python intake.py --input X --kb-root K --out-plan plan.json

    # 3) 人工修正后执行（decisions.json 见下方格式）
    python intake.py --input X --kb-root K --decisions d.json --apply

    # 4) 高置信直接落盘（慎用，仅用于完全可信的批量补录）
    python intake.py --input X --kb-root K --apply --auto

decisions.json 格式（键支持"文件名"或"路径片段"）：
    {
      "20250815_报水利部.docx": {
        "action": "accept",
        "target_dir": "01-龚利民（董事长）/对外汇报（部委）",
        "note": "人工确认，密级敏感但目录性质一致"
      },
      "某不确定文件.docx": {"action": "skip", "note": "内容存疑，暂不入库"}
    }

退出码：0 成功（含"有阻塞项待确认"）；2 参数/资产错误；3 落盘过程有失败项。
"""
from __future__ import annotations

import argparse
import io
import json
import re
import shutil
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path, PurePosixPath
from typing import Any

_SCRIPT_DIR = Path(__file__).resolve().parent
if str(_SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPT_DIR))

if sys.stdout.encoding is None or sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass

try:
    import classify as C
    import state as S
except ImportError as e:  # pragma: no cover
    sys.stderr.write(f"[FAIL] 无法导入同目录脚本：{e}\n")
    raise SystemExit(2)


# ---------------------------------------------------------------------------
# 计划条目
# ---------------------------------------------------------------------------


@dataclass
class PlanItem:
    seq: int
    action: str = "block"            # classify 给出的分流：silent/preview/block
    decision: str = "deferred"       # accepted / rejected / deferred / skipped
    src: str = ""
    md5: str = ""
    size: int = 0
    doc_type: str | None = None
    subject: str | None = None
    audience: str | None = None
    security: str | None = None
    date: str = ""
    title_clean: str = ""
    tags: list[str] = field(default_factory=list)
    target_dir: str | None = None
    new_filename: str = ""
    dest_path: str = ""
    confidence: float = 0.0
    evidence: list[str] = field(default_factory=list)
    blockers: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    candidates: list[dict] = field(default_factory=list)
    reason: str = ""                 # 为什么是这个 decision
    overridden: list[str] = field(default_factory=list)   # 被人为改过的字段
    error: str = ""


# ---------------------------------------------------------------------------
# 辅助
# ---------------------------------------------------------------------------


def sanitize_filename(name: str, forbidden: list[str]) -> str:
    """剔除文件名非法字符。注意 Windows 上还禁止尾随空格与点。"""
    out = name
    for ch in forbidden:
        out = out.replace(ch, "-")
    out = re.sub(r"[\x00-\x1f]", "", out)
    out = out.rstrip(" .")
    return out or "未命名"


def resolve_decision(decisions: dict, src: Path) -> dict | None:
    """按 文件名 → 相对路径片段 → 绝对路径 三级匹配人工决定。"""
    if not decisions:
        return None
    cands = [src.name, src.stem, str(src), src.as_posix()]
    for key in cands:
        if key in decisions:
            return decisions[key]
    # 子串匹配（容错：用户只记得半截名字）
    for key, val in decisions.items():
        if key and (key in src.name or key in src.as_posix()):
            return val
    return None


def rebuild_filename(item: PlanItem, assets: C.Assets) -> str:
    """按 taxonomy 的命名模板重新拼文件名（供字段被人工覆盖后刷新）。"""
    meta = assets.taxonomy["meta"]
    subj = meta["subject_placeholder"] if item.subject in (None, "-") else item.subject
    d = (item.date or "").replace("-", "")
    return f"{d}_{item.doc_type or '未定'}_{item.title_clean or '未命名'}_{subj}{Path(item.src).suffix.lower()}"


# ---------------------------------------------------------------------------
# 计划构建
# ---------------------------------------------------------------------------


def build_plan(files: list[Path], assets: C.Assets, kb_root: Path,
               st: S.StateStore, decisions: dict,
               auto: bool, accept_preview: bool) -> list[PlanItem]:
    forbidden = assets.taxonomy["meta"].get("forbidden_chars", [])
    items: list[PlanItem] = []

    for i, f in enumerate(files, 1):
        try:
            md5 = S.md5_file(f)
        except Exception as e:
            items.append(PlanItem(seq=i, src=str(f), error=f"读取失败：{e}",
                                  decision="rejected", reason="源文件不可读"))
            continue

        it = PlanItem(seq=i, src=str(f), md5=md5, size=f.stat().st_size)

        # ---- 幂等闸门：内容已进过库就跳过 ----
        if st.seen(md5):
            prev = st.get(md5) or {}
            it.decision = "skipped"
            it.action = "silent"
            # 回填历史判定，让"跳过"这一行也能读懂，而不是一片"未判定"
            it.doc_type = prev.get("doc_type")
            it.subject = prev.get("subject")
            it.audience = prev.get("audience")
            it.security = prev.get("security")
            it.date = prev.get("date") or ""
            it.title_clean = prev.get("title_clean") or ""
            it.tags = list(prev.get("tags") or [])
            it.confidence = float(prev.get("confidence") or 0.0)
            rel = prev.get("target_rel") or ""
            if rel:
                rel_p = PurePosixPath(rel)          # target_rel 是 posix 风格，且【已含文件名】
                it.target_dir = "" if str(rel_p.parent) == "." else str(rel_p.parent)
                it.new_filename = rel_p.name
                it.dest_path = str(kb_root.joinpath(*rel_p.parts))
            it.reason = f"内容已在库中（{prev.get('updated','?')} 登记），跳过"
            items.append(it)
            continue

        v = C.classify_one(f, assets, kb_root, use_content=True)

        it.action = v.action
        it.doc_type = v.doc_type
        it.subject = v.subject
        it.audience = v.audience
        it.security = v.security
        it.tags = list(v.tags)
        it.date = v.date
        it.title_clean = v.title_clean
        it.target_dir = v.target_dir
        it.new_filename = v.new_filename or ""
        it.confidence = v.confidence
        it.evidence = v.evidence
        it.blockers = list(v.blockers)
        it.notes = list(v.notes)
        it.candidates = v.candidates

        # ---- 人工决定优先于一切 ----
        dec = resolve_decision(decisions, f)
        if dec is not None:
            act = (dec.get("action") or "").strip().lower()
            if act in ("skip", "rejected", "reject"):
                it.decision = "rejected"
                it.reason = dec.get("note") or "人工决定：不入库"
                items.append(it)
                continue
            if act in ("accept", "accepted", "ok"):
                for fld in ("doc_type", "subject", "audience", "security", "date", "title_clean"):
                    if dec.get(fld) not in (None, ""):
                        setattr(it, fld, dec[fld])
                        it.overridden.append(fld)
                if dec.get("target_dir"):
                    it.target_dir = dec["target_dir"]
                    it.overridden.append("target_dir")
                if dec.get("new_filename"):
                    it.new_filename = dec["new_filename"]
                    it.overridden.append("new_filename")
                elif it.overridden:
                    it.new_filename = rebuild_filename(it, assets)
                it.decision = "accepted"
                it.reason = dec.get("note") or "人工确认接受（推翻引擎判定）"
                it.blockers = []
                items.append(it)
                continue
            # 未识别的 action → 退回默认流程，但记一笔
            it.notes.append(f"[~] decisions 里的 action「{act}」无法识别，按默认流程处理")

        # ---- 默认流程 ----
        if not it.target_dir:
            it.decision = "deferred"
            it.reason = "未能确定目标目录（缺少主空间或类型映射），需人工指定"
        elif it.action == "silent":
            it.decision = "accepted" if auto else "deferred"
            it.reason = "高置信，自动接受" if auto else "高置信，等待 --auto 或人工确认"
        elif it.action == "preview":
            it.decision = "accepted" if (auto or accept_preview) else "deferred"
            it.reason = "中置信，已确认接受" if (auto or accept_preview) else "中置信，建议人工扫一眼再放行"
        else:
            it.decision = "deferred"
            it.reason = "低置信/撞护栏，必须人工定夺"

        # ---- 护栏硬拦截：blocker 存在时，任何开关都不能放行 ----
        if it.blockers:
            it.decision = "deferred"
            it.reason = "存在阻塞项：" + it.blockers[0]

        # ---- 目标路径 + 命名规整 ----
        if it.target_dir:
            it.new_filename = sanitize_filename(it.new_filename, forbidden)
            it.dest_path = str(kb_root / it.target_dir / it.new_filename)

        # ---- 目标冲突检测 ----
        if it.decision == "accepted" and it.dest_path:
            dest = Path(it.dest_path)
            if dest.exists():
                try:
                    dmd5 = S.md5_file(dest)
                except Exception:
                    dmd5 = ""
                if dmd5 == md5:
                    it.decision = "skipped"
                    it.reason = "目标已存在且内容一致，幂等复用"
                else:
                    it.decision = "deferred"
                    it.blockers.append(
                        f"目标已存在且内容不同：{it.target_dir}/{it.new_filename} "
                        "—— 禁止自动覆盖，请改名或人工处理"
                    )
                    it.reason = "目标命名冲突"

        items.append(it)

    return items


# ---------------------------------------------------------------------------
# 执行
# ---------------------------------------------------------------------------


def apply_plan(items: list[PlanItem], kb_root: Path,
               st: S.StateStore, log: S.IntakeLog, batch_id: str,
               operator: str) -> tuple[int, int]:
    """执行 accepted 项。返回 (成功数, 失败数)。"""
    ok = fail = 0
    rows: list[dict] = []

    for it in items:
        if it.decision != "accepted":
            rows.append(_log_row(it, batch_id, operator, decision=it.decision))
            continue

        src, dest = Path(it.src), Path(it.dest_path)
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            # copy2 保留 mtime —— 知识库里"文件时间"有时是唯一线索
            shutil.copy2(src, dest)
            # 校验落地完整性
            if S.md5_file(dest) != it.md5:
                raise RuntimeError("复制后 md5 不一致")

            it.reason = (it.reason + " · 已落盘").strip(" ·")
            st.put(it.md5,
                   target_rel=str(dest.relative_to(kb_root)).replace("\\", "/"),
                   new_filename=dest.name,
                   doc_type=it.doc_type, subject=it.subject,
                   audience=it.audience, security=it.security,
                   date=it.date, title_clean=it.title_clean, tags=it.tags,
                   confidence=it.confidence,
                   batch_id=batch_id, src=str(src), size=it.size,
                   status="local")
            ok += 1
        except Exception as e:
            it.error = str(e)
            it.decision = "deferred"
            it.blockers.append(f"落盘失败：{e}")
            fail += 1
        rows.append(_log_row(it, batch_id, operator, decision=it.decision))

    if rows:
        log.append(rows)
    if ok or fail:
        st.save()
    return ok, fail


def _log_row(it: PlanItem, batch_id: str, operator: str, decision: str) -> dict:
    return {
        "ts": S.now_iso(),
        "batch_id": batch_id,
        "action": it.action,
        "decision": decision,
        "src_path": it.src,
        "src_size": it.size,
        "md5": it.md5,
        "doc_type": it.doc_type or "",
        "subject": it.subject or "",
        "audience": it.audience or "",
        "security": it.security or "",
        "date": it.date,
        "title_clean": it.title_clean,
        "target_dir": it.target_dir or "",
        "new_filename": it.new_filename,
        "confidence": it.confidence,
        "dest_path": it.dest_path if decision == "accepted" else "",
        "operator": operator,
        "notes": " | ".join(
            filter(None, [it.reason, it.error,
                          ("人工覆盖：" + ",".join(it.overridden)) if it.overridden else "",
                          *it.blockers])
        )[:900],
    }


# ---------------------------------------------------------------------------
# 渲染
# ---------------------------------------------------------------------------

_ICON = {"accepted": "[+] 入库", "skipped": "[=] 跳过", "deferred": "[?] 待定", "rejected": "[x] 不入库"}
_ACT = {"silent": "高", "preview": "中", "block": "低"}


def fmt_item(it: PlanItem) -> str:
    lines = [
        f"  {_ICON.get(it.decision, it.decision)}  #{it.seq}  {Path(it.src).name}",
        f"      {it.doc_type or '未判定'} · {it.subject or '未判定'} · "
        f"{it.audience or '-'} · {it.security or '-'} · {it.date} · 置信 {it.confidence:.2f}({_ACT.get(it.action,'?')})",
    ]
    if it.target_dir:
        lines.append(f"      → {it.target_dir}/{it.new_filename}")
    if it.reason:
        lines.append(f"      理由：{it.reason}")
    for b in it.blockers:
        lines.append(f"      [!] {b}")
    for c in it.candidates[:2]:
        lines.append(f"      候选：{c.get('sub_dir')} ← {c.get('reason')}")
    if it.error:
        lines.append(f"      [x] {it.error}")
    return "\n".join(lines)


def print_plan(items: list[PlanItem], kb_root: Path, batch_id: str, applied: bool) -> None:
    counts: dict[str, int] = {}
    for it in items:
        counts[it.decision] = counts.get(it.decision, 0) + 1

    print("=" * 78)
    print(f"知识库增量接入{'执行结果' if applied else '计划（dry-run，未写盘）'}  批次 {batch_id}")
    print("=" * 78)
    print(f"  知识库根  {kb_root}")
    print(f"  共 {len(items)} 份：  " + " · ".join(
        f"{_ICON.get(k, k)} {v}" for k, v in sorted(counts.items())))
    print()

    for grp, title in (("accepted", "将入库"), ("skipped", "跳过（幂等）"),
                       ("deferred", "待人工确认"), ("rejected", "不入库")):
        subset = [it for it in items if it.decision == grp]
        if not subset:
            continue
        print(f"{'-' * 78}\n【{title}】{len(subset)} 份\n")
        for it in subset:
            print(fmt_item(it))
            print()

    if not applied:
        pend = counts.get("deferred", 0)
        print("=" * 78)
        print("下一步：")
        if counts.get("accepted"):
            print("  · 计划中的高置信项已就绪 → 加 --apply 落盘")
        if pend:
            print(f"  · 有 {pend} 份待定 → 用 --decisions decisions.json 逐项确认后重跑")
        print("  · 复看完整计划 → --out-plan plan.json")


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------


def collect_files(target: Path) -> list[Path]:
    if target.is_dir():
        return sorted(
            p for p in target.rglob("*")
            if p.is_file() and p.suffix.lower() in C.SUPPORTED_EXT
            and not p.name.startswith("~$") and not p.name.startswith(".")
        )
    if target.is_file():
        return [target]
    return []


def main() -> int:
    ap = argparse.ArgumentParser(
        description="知识库对话式增量接入（默认 dry-run，只复制不移动）",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", required=True, help="新文件或目录")
    ap.add_argument("--kb-root", required=True, help="知识库根目录")
    ap.add_argument("--skill-dir", help="skill 根目录（默认取本脚本上级）")
    ap.add_argument("--decisions", help="人工决定 JSON 文件")
    ap.add_argument("--apply", action="store_true", help="真正落盘（不加则只出计划）")
    ap.add_argument("--auto", action="store_true", help="连中置信也一并接受（仍受护栏约束）")
    ap.add_argument("--accept-preview", action="store_true", help="仅接受中置信项，不碰低置信")
    ap.add_argument("--out-plan", help="把计划写入 JSON")
    ap.add_argument("--operator", default="agent", help="决策者标记（auto/user/agent）")
    ap.add_argument("--no-content", action="store_true", help="只看文件名，不解析正文")
    args = ap.parse_args()

    skill_dir = Path(args.skill_dir) if args.skill_dir else _SCRIPT_DIR.parent
    assets = C.Assets(skill_dir)
    kb_root = Path(args.kb_root).resolve()
    if not kb_root.exists():
        sys.stderr.write(f"[FAIL] 知识库根不存在：{kb_root}\n")
        return 2

    files = collect_files(Path(args.input))
    if not files:
        sys.stderr.write(f"[FAIL] 没找到可处理的文档：{args.input}\n")
        return 2

    decisions: dict = {}
    if args.decisions:
        dp = Path(args.decisions)
        if not dp.exists():
            sys.stderr.write(f"[FAIL] decisions 文件不存在：{dp}\n")
            return 2
        decisions = json.loads(dp.read_text(encoding="utf-8"))

    st = S.StateStore(kb_root)
    log = S.IntakeLog(kb_root)
    batch_id = S.new_batch_id()

    # --no-content 时退化为纯文件名判定：临时把抽取函数短路
    if args.no_content:
        C.extract_text = lambda p: C.DocText("", "skipped", False, note="--no-content 已跳过正文解析")

    items = build_plan(files, assets, kb_root, st, decisions,
                       auto=args.auto, accept_preview=args.accept_preview)

    if args.out_plan:
        payload = {
            "batch_id": batch_id,
            "kb_root": str(kb_root),
            "generated_at": S.now_iso(),
            "summary": {k: sum(1 for i in items if i.decision == k)
                        for k in ("accepted", "skipped", "deferred", "rejected")},
            "items": [asdict(i) for i in items],
        }
        S.atomic_write_text(Path(args.out_plan),
                            json.dumps(payload, ensure_ascii=False, indent=2))

    if not args.apply:
        print_plan(items, kb_root, batch_id, applied=False)
        # dry-run 不写日志（日志是"事实记录"，没执行就没事实）
        if args.out_plan:
            print(f"\n计划 JSON 已写入：{args.out_plan}")
        return 0

    operator = "auto" if args.auto else args.operator
    ok, fail = apply_plan(items, kb_root, st, log, batch_id, operator)
    print_plan(items, kb_root, batch_id, applied=True)
    print()
    print(f"落盘完成：成功 {ok} · 失败 {fail}")
    print(f"台账已更新：{S.STATE_NAME} / {S.LOG_NAME}")
    return 3 if fail else 0


if __name__ == "__main__":
    raise SystemExit(main())
