#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""state.py —— 知识库热更新的【台账层】

三本账，各管一件事，不得混淆：

    _intake_log.csv   入库审计（append-only，人可读，Excel 可开）
                      「谁、什么时候、把哪个文件、按什么判定、放到了哪里、成没成」

    _sync_state.json  幂等台账（可覆盖，机可读）
                      「这个内容的 md5 是否已经进过库 / 是否已经同步到乐享」

    _sync_log.csv     同步审计（append-only，人可读）
                      「哪份内容、什么时候、以什么动作、推到了乐享的哪个条目」

为什么要分开：
  * 日志是**事实记录**，只增不改 —— 出了问题要能回溯，不能因为重跑就被抹掉
  * 状态是**当前视图**，必须可覆盖 —— 幂等判断只看"现在是什么样"
  * 入库审计与同步审计回答的是**两个不同的问题**：
        「这份文件该不该进库、该放哪」  ↔  「这份内容到底有没有推上去」
    混在一张表里，两边的统计都会被对方污染

放在知识库根目录下，以 `_` 开头，天然被 classify.py 的扩展名白名单过滤掉，
不会被当成待分类文档。
"""
from __future__ import annotations

import csv
import hashlib
import io
import json
import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

if sys.stdout.encoding is None or sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass

LOG_NAME = "_intake_log.csv"
SYNC_LOG_NAME = "_sync_log.csv"
STATE_NAME = "_sync_state.json"
STATE_VERSION = 1
AUTH_NAME = "_authorized_domains.json"

# 审计日志的列定义 —— 顺序即 CSV 列序，改动需同步 migration
LOG_COLUMNS = [
    "ts",             # 处理时间（本地时区）
    "batch_id",       # 批次号，一次 intake 调用一个
    "action",         # 判定动作：silent / preview / block
    "decision",       # 最终决定：accepted / rejected / deferred / skipped
    "src_path",       # 原始文件绝对路径
    "src_size",       # 原始字节数
    "md5",            # 内容指纹（幂等键）
    "doc_type",       # 文档类型
    "subject",        # 主体/领导
    "audience",       # 受众
    "security",       # 密级
    "date",           # 文档日期 YYYY-MM-DD
    "title_clean",    # 清理后标题
    "target_dir",     # 目标目录（相对知识库根）
    "new_filename",   # 目标文件名
    "confidence",     # 置信度
    "dest_path",      # 实际落地路径（未执行时为空）
    "operator",       # 决策者：auto / user / agent
    "notes",          # 备注（冲突、降级、人工修正等）
]

# 同步日志的列定义 —— 记录"内容 → 乐享条目"的传输事实
SYNC_LOG_COLUMNS = [
    "ts",               # 处理时间
    "batch_id",         # 同步批次号
    "md5",              # 内容指纹（对应 _sync_state.json 的主键）
    "doc_type",         # 文档类型（便于按类型统计成功率）
    "subject",          # 主体/领导
    "target_rel",       # 本地相对路径（posix）
    "remote_folder",    # 远端目录链
    "action",           # 动作：create_folder / upload / reuse / skip / error
    "remote_entry_id",  # 远端条目 id
    "remote_url",       # 远端链接
    "size",             # 字节数
    "elapsed_s",        # 该条耗时（秒）
    "tags",             # 写入的标签（逗号分隔）
    "operator",         # 执行者
    "notes",            # 备注 / 错误信息
]


# ---------------------------------------------------------------------------
# 工具
# ---------------------------------------------------------------------------


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def md5_file(path: Path, chunk: int = 1 << 20) -> str:
    """分块计算 md5 —— 知识库里 400MB+ 的 pptx 不少，不能一次读进内存。"""
    h = hashlib.md5()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


def md5_bytes(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()


def new_batch_id() -> str:
    return "B" + datetime.now().strftime("%Y%m%d%H%M%S")


def atomic_write_text(path: Path, text: str, encoding: str = "utf-8") -> None:
    """先写临时文件再 os.replace —— 中途崩溃不会留下半个 JSON。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix=".tmp_", suffix=path.suffix)
    try:
        with os.fdopen(fd, "w", encoding=encoding, newline="") as f:
            f.write(text)
        os.replace(tmp, path)
    except Exception:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


# ---------------------------------------------------------------------------
# 幂等台账 _sync_state.json
# ---------------------------------------------------------------------------


class StateStore:
    """幂等台账。

    以 **内容 md5** 为主键 —— 这是唯一稳定、与路径无关的身份。
    文件改名、换目录都不影响"这份内容我见过没有"的判断。

    另外维护一条 `by_target` 反向索引（target_rel → md5），用来发现
    「同一个目标路径被两份不同内容争抢」的冲突。
    """

    def __init__(self, kb_root: Path, filename: str = STATE_NAME):
        self.root = Path(kb_root)
        self.path = self.root / filename
        self.data: dict[str, Any] = {
            "version": STATE_VERSION,
            "updated": "",
            "files": {},       # md5 -> record
        }
        self.load()

    # ---- 读写 ----

    def load(self) -> None:
        if not self.path.exists():
            return
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
            if raw.get("version") != STATE_VERSION:
                # 版本不匹配：不猜测语义，保留原文件另存备份，从空表重来
                bak = self.path.with_suffix(f".v{raw.get('version')}.bak.json")
                if not bak.exists():
                    bak.write_text(json.dumps(raw, ensure_ascii=False, indent=2),
                                   encoding="utf-8")
                return
            self.data = raw
        except Exception as e:
            sys.stderr.write(f"[WARN] {self.path.name} 读取失败（按空台账继续）：{e}\n")

    def save(self) -> None:
        self.data["version"] = STATE_VERSION
        self.data["updated"] = now_iso()
        atomic_write_text(self.path, json.dumps(self.data, ensure_ascii=False, indent=2))

    # ---- 查询 ----

    @property
    def files(self) -> dict[str, dict]:
        return self.data.setdefault("files", {})

    def get(self, md5: str) -> dict | None:
        return self.files.get(md5)

    def seen(self, md5: str) -> bool:
        """内容是否曾经入库（不论成败）。"""
        return md5 in self.files

    def is_synced(self, md5: str) -> bool:
        """内容是否已成功同步到乐享。"""
        rec = self.files.get(md5)
        return bool(rec and rec.get("remote_entry_id"))

    def by_target(self, target_rel: str) -> dict | None:
        """按目标相对路径反查 —— 用于发现同名冲突。"""
        for md5, rec in self.files.items():
            if rec.get("target_rel") == target_rel:
                return {"md5": md5, **rec}
        return None

    # ---- 写入 ----

    def put(self, md5: str, **fields) -> dict:
        rec = self.files.setdefault(md5, {})
        rec.update(fields)
        rec["updated"] = now_iso()
        return rec

    def drop(self, md5: str) -> None:
        self.files.pop(md5, None)


# ---------------------------------------------------------------------------
# CSV 审计日志基类
# ---------------------------------------------------------------------------


class _CsvLog:
    """append-only CSV 台账的公共实现。

    编码用 **utf-8-sig**（带 BOM）—— 因为这两本账的目的之一就是给人用 Excel 看，
    没有 BOM 的 UTF-8 CSV 在中文 Windows 上会乱码。
    """

    columns: list[str] = []
    filename: str = ""

    def __init__(self, kb_root: Path, filename: str | None = None):
        self.root = Path(kb_root)
        self.path = self.root / (filename or self.filename)

    def append(self, rows: Iterable[dict]) -> int:
        rows = list(rows)
        if not rows:
            return 0
        self.root.mkdir(parents=True, exist_ok=True)
        is_new = not self.path.exists()
        with open(self.path, "a", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=self.columns, extrasaction="ignore")
            if is_new:
                w.writeheader()
            for r in rows:
                w.writerow({c: r.get(c, "") for c in self.columns})
        return len(rows)

    def read_all(self) -> list[dict]:
        if not self.path.exists():
            return []
        with open(self.path, "r", encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f))

    def batches(self, batch_id: str) -> list[dict]:
        return [r for r in self.read_all() if r.get("batch_id") == batch_id]

    def last_batch_id(self) -> str | None:
        rows = self.read_all()
        return rows[-1]["batch_id"] if rows else None

    def count(self) -> int:
        return len(self.read_all())


# ---------------------------------------------------------------------------
# 入库审计 _intake_log.csv
# ---------------------------------------------------------------------------


class IntakeLog(_CsvLog):
    """入库决策的事实记录。"""

    columns = LOG_COLUMNS
    filename = LOG_NAME


# ---------------------------------------------------------------------------
# 同步审计 _sync_log.csv
# ---------------------------------------------------------------------------


class SyncLog(_CsvLog):
    """向乐享推送的传输事实记录。"""

    columns = SYNC_LOG_COLUMNS
    filename = SYNC_LOG_NAME


# ---------------------------------------------------------------------------
# CLI：查看台账（只读，便于对话中自查）
# ---------------------------------------------------------------------------


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser(description="知识库热更新台账（只读查看）")
    ap.add_argument("--kb-root", required=True, help="知识库根目录")
    ap.add_argument("--show", choices=["summary", "state", "log", "sync"], default="summary")
    ap.add_argument("--limit", type=int, default=30)
    args = ap.parse_args()

    kb = Path(args.kb_root)
    st = StateStore(kb)
    lg = IntakeLog(kb)

    if args.show == "state":
        print(json.dumps(st.data, ensure_ascii=False, indent=2))
        return 0

    if args.show == "log":
        rows = lg.read_all()
        for r in rows[-args.limit:]:
            print(f"{r.get('ts','')}  {r.get('decision',''):<9} "
                  f"{r.get('new_filename','') or r.get('src_path','')}")
        print(f"\n共 {len(rows)} 条记录")
        return 0

    if args.show == "sync":
        rows = SyncLog(kb).read_all()
        for r in rows[-args.limit:]:
            print(f"{r.get('ts','')}  {r.get('action',''):<14} "
                  f"{r.get('remote_entry_id',''):<22} {r.get('target_rel','')}")
        print(f"\n共 {len(rows)} 条记录")
        return 0

    rows = lg.read_all()
    files = st.files
    synced = sum(1 for r in files.values() if r.get("remote_entry_id"))
    sync_rows = SyncLog(kb).count()
    print("=" * 68)
    print("知识库热更新台账")
    print("=" * 68)
    print(f"  知识库根      {kb}")
    print(f"  台账文件      {LOG_NAME} / {STATE_NAME} / {SYNC_LOG_NAME}")
    print(f"  已登记内容    {len(files)} 条（其中已同步乐享 {synced} 条）")
    print(f"  入库审计      {len(rows)} 行")
    print(f"  同步审计      {sync_rows} 行")
    if files:
        print(f"  最近更新      {st.data.get('updated', '-')}")
    if rows:
        print(f"  最近批次      {lg.last_batch_id()}")
        counts: dict[str, int] = {}
        for r in rows:
            counts[r.get("decision", "?")] = counts.get(r.get("decision", "?"), 0) + 1
        print("  决定分布      " + " · ".join(f"{k} {v}" for k, v in sorted(counts.items())))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
