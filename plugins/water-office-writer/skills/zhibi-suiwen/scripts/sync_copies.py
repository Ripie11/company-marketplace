#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sync_copies.py —— 文秘专员（water-office-writer）跨副本一致性巡检与同步。

分三层管辖：

| 层 | 副本 | 内容 |
|----|------|------|
| **skill 层** | **P / Q / R** 三副本 × N 个技能 | 各技能的 `SKILL.md` / `assets/` / `scripts/` / `references/` / `selftest/` |
| **插件层** | **P / R** 两副本 | `plugin.json` / `README.md` / `USER_GUIDE.md` / `agents/` / `avatars/` |

## 三副本

| 代号 | 位置 | 角色 |
|------|------|------|
| **P** | `~/.workbuddy/plugins/marketplaces/my-experts/plugins/water-office-writer/skills/<skill>` | 插件包（默认规范源） |
| **Q** | `<workspace>/.workbuddy/skills/<skill>` | 项目级 skill |
| **R** | `<workspace>/专家分发/company-marketplace/plugins/water-office-writer/skills/<skill>` | 专家分发包 |

插件层只是 P ↔ R 两个副本 —— Q 是裸 skill 目录，没有插件外壳，不参与。

## 管辖哪些技能

本工具按 `SKILLS` 注册表逐个技能巡检。**新增技能必须在此登记**，否则它的漂移
永远不会被发现（2026-09-14 实际踩过：`knowledge-base-organizer` 三副本漂移了
4 个文件，直到人工统一"总裁"称谓时才被翻出来）。

## 规范源与例外

skill 层默认规范源 = **P**。少数文件因 P 侧为旧版/缺失，在各技能的 `TAKE_FROM`
中显式声明取自他副本 —— **这些表是跨副本版本分叉的唯一台账**。

当前台账：

| 技能 | 文件 | 取自 | 原因 |
|------|------|------|------|
| zhibi-suiwen | `scripts/build_style_profile.py` | Q | v3 多源分通道 + format_hints |
| zhibi-suiwen | `scripts/extract_md.py` | Q | P 侧缺失 |
| zhibi-suiwen | `scripts/extract_docx.py` | R | `line_spacing.pt is not None` 空值护栏 |
| zhibi-suiwen | `scripts/extract_pdf.py` | R | v3 OCR + 坐标级段落重构 |
| knowledge-base-organizer | `assets/taxonomy.yaml` | R | 「方琳（总裁）」改称后的正名版 |
| knowledge-base-organizer | `scripts/backtest.py` | R | 同上（含旧名→新名迁移映射） |
| knowledge-base-organizer | `scripts/kbsync.py` | R | 同上（远端乐享已统一为 `02-方琳（总裁）`） |
| knowledge-base-organizer | `scripts/reorganize_template.py` | R | 同上 |

若某个文件在 Q / R 侧更新而 P 侧仍是旧版，**必须**把它登记进对应 `TAKE_FROM`，
否则会被 P 侧旧版反向覆盖。`--check` 会打印每个漂移项的当前哈希便于判断。

插件层同理，台账是 `PLUGIN_TAKE_FROM`（裁决依据见该常量上方的注释）。
已知有差异但**尚未裁决**的文件放进 `PLUGIN_UNRESOLVED`，只报告、不自动同步
—— 绝不把猜测当事实写进副本。

## 身份替换

仅 `zhibi-suiwen` 需要改身份：P / R 的技能名是 `zhibi-suiwen`，Q 是
`leadership-report-generator`，文本类文件同步时按目标副本替换（含 `# 标题` 行、
frontmatter `name`、档案 `generated_by`、`selftest/_manifest.json` 内的绝对路径）。

`knowledge-base-organizer` 三副本同名，**不做替换**，纯字节拷贝。
插件层两侧技能名相同，也不做替换。

两处保护：
- `NO_TRANSFORM`：列出的文件不做身份替换（自身必须指名道姓写出"另一个副本叫什么"）
- 源文本**两个名字并存**时自动跳过整份替换并告警 —— 宁可原样，绝不猜错

## 字节级 I/O 铁律

所有读写一律走 `read_bytes` / `write_bytes`，**绝不** `Path.read_text` / `write_text`。

原因（2026-09-14 实际踩坑，代价是 15 个文件被破坏）：`write_text` 在 Windows 上把
`\n` 翻译成 `\r\n`；若待写文本里**已经**是 CRLF，就会写出 `\r\r\n`。而
`str.splitlines()` 把 `\r\r\n` 视为**两个**换行 → Markdown 表格行之间多出空行 →
`md_to_gongwen_docx.py` 的 `parse_md_table`（要求表头下一行即分隔行）判定失败 →
表格整块退化成普通段落，落版静默失效。

`unfold_cr()` 在**读取端**折叠 `\r\r\n` → `\r\n`：腐坏是"多副本同时中招"，
从任一副本派生的期望值都会继承腐坏，只有读取端折叠才能让 `--apply` 真正自愈。
`verify_copy()` 会对各副本逐文件复扫，`--check` 亦包含。

## 用法

    python sync_copies.py --check                     # 只巡检（默认），有漂移退出码 1
    python sync_copies.py --apply                     # 执行同步（skill 层 + 插件层）
    python sync_copies.py --apply --only scripts      # 只同步某几个顶层目录
    python sync_copies.py --apply --skill zhibi-suiwen  # 只处理某个技能
    python sync_copies.py --apply --no-verify         # 跳过同步后的语法校验
    python sync_copies.py --apply --no-plugin         # 只处理 skill 层，不碰插件层
    python sync_copies.py --prune                     # 顺带删除 *.bak / __pycache__

## 退出码

`--check` 下发现漂移或行尾腐坏 → 1；一致 → 0。可直接用于改动前后自检 / CI。
`PLUGIN_UNRESOLVED` 中待裁决的差异**不影响退出码**（已知且经人工确认的差异）。

生命周期：**常驻工具**，幂等可重复运行。与 `_build/reorganize.py` 那类一次性
冷启动脚本不同，本脚本除 `--prune` 外不做任何删除动作。
"""
from __future__ import annotations

import argparse
import ast
import hashlib
import json
import os
import shutil
import sys
from pathlib import Path

# ---------------------------------------------------------------- 常量

NAME_P = "zhibi-suiwen"
NAME_Q = "leadership-report-generator"
TITLE_P = "# zhibi-suiwen — 执笔随文"
TITLE_Q = "# leadership-report-generator — 领导风格报告生成器"

# 文本类扩展名：同步时做身份替换
TEXT_EXT = {".md", ".yaml", ".yml", ".txt", ".py", ".json", ".csv"}

# 各技能通用忽略（目录 / 后缀）
IGNORE_DIRS = {"_backup_v1", "__pycache__", ".pytest_cache"}
IGNORE_SUFFIX = {".bak", ".pyc", ".pyo"}


# ---------------------------------------------------------------- 技能注册表
#
# ⚠️ 新增技能必须登记在这里，否则它的三副本漂移永远不会被发现。
# 每个技能一组配置：根路径、版本分叉台账、身份替换、忽略规则。

class SkillCfg:
    """单个 skill 的同步配置。"""

    def __init__(self, key: str, label: str, roots_fn,
                 take_from: dict[str, str] | None = None,
                 path_rewrite_rel: str | None = None,
                 no_transform: frozenset[str] = frozenset(),
                 ignore_relprefix: tuple[str, ...] = (),
                 ignore_relnames: frozenset[str] = frozenset(),
                 rename: bool = False):
        self.key = key
        self.label = label
        self._roots_fn = roots_fn          # callable(workspace) -> {code: Path}
        self.roots: dict[str, Path] = {}   # 由 resolve_roots() 物化
        self.take_from = take_from or {}
        self.path_rewrite_rel = path_rewrite_rel
        self.no_transform = no_transform
        self.ignore_relprefix = ignore_relprefix
        self.ignore_relnames = ignore_relnames
        # rename=True 时按目标副本替换技能名（仅 zhibi-suiwen 需要）
        self.rename = rename

    def resolve_roots(self, workspace: Path) -> None:
        """按工作区根物化三副本路径（main 里调用一次）。"""
        self.roots = self._roots_fn(workspace)


def _zs_roots(workspace: Path) -> dict[str, Path]:
    base = Path.home() / ".workbuddy/plugins/marketplaces/my-experts/plugins/water-office-writer/skills"
    return {
        "P": base / "zhibi-suiwen",
        "Q": workspace / ".workbuddy/skills/leadership-report-generator",
        "R": workspace / "专家分发/company-marketplace/plugins/"
                         "water-office-writer/skills/zhibi-suiwen",
    }


def _kbo_roots(workspace: Path) -> dict[str, Path]:
    base = Path.home() / ".workbuddy/plugins/marketplaces/my-experts/plugins/water-office-writer/skills"
    return {
        "P": base / "knowledge-base-organizer",
        "Q": workspace / ".workbuddy/skills/knowledge-base-organizer",
        "R": workspace / "专家分发/company-marketplace/plugins/"
                         "water-office-writer/skills/knowledge-base-organizer",
    }


SKILLS: list[SkillCfg] = [
    SkillCfg(
        key="zhibi-suiwen",
        label="执笔随文（写得出）",
        roots_fn=_zs_roots,
        take_from={
            "scripts/build_style_profile.py": "Q",
            "scripts/extract_md.py": "Q",
            "scripts/extract_docx.py": "R",
            "scripts/extract_pdf.py": "R",
        },
        path_rewrite_rel="selftest/cases/_manifest.json",
        # 这两个文件天生必须指名道姓写出"另一个副本叫什么"（脚本自身硬编码 P/Q/R
        # 路径与 NAME_P/NAME_Q 常量；selftest/README.md 的「三副本一致性」表说明
        # "项目级 skill 的技能名是 X"），故整份原样拷贝。
        # ⚠️ 连带约定：其余文件正文禁止硬写具体技能名，用「本技能」中性措辞，
        #    否则触发"两名并存 → 跳过整份替换"护栏（见 SKILL.md §十 铁律 3）。
        no_transform=frozenset({"scripts/sync_copies.py", "selftest/README.md"}),
        ignore_relprefix=("selftest/out_v1/", "selftest/out_v2/"),
        ignore_relnames=frozenset({
            "selftest/verify.json",
            "selftest/v1_verify.json",
            "selftest/v2_verify.json",
        }),
        rename=True,
    ),
    SkillCfg(
        key="knowledge-base-organizer",
        label="知识库整理（找得到）",
        roots_fn=_kbo_roots,
        # 「方琳（总裁）」改称（2026-09-14 用户确认「执行总裁」为误称）：
        # 远端乐享主空间已统一为 `02-方琳（总裁）`，R 侧 4 个文件是改后权威版，
        # P / Q 侧仍是旧版 → 台账声明取 R。旧名仅在 aliases / 迁移映射中保留
        # （向后兼容已归档素材），这是 R 版本有意为之，不要"顺手清理"。
        take_from={
            "assets/taxonomy.yaml": "R",
            "scripts/backtest.py": "R",
            "scripts/kbsync.py": "R",
            "scripts/reorganize_template.py": "R",
        },
        # 三副本同名，不做身份替换
        rename=False,
    ),
]


# ---------------------------------------------------------------- 定位

def find_workspace() -> Path:
    """定位工作区根：$WATER_WORKSPACE > 自 CWD 向上找任一项目级 skill。"""
    env = os.environ.get("WATER_WORKSPACE")
    if env:
        return Path(env).resolve()
    cur = Path.cwd().resolve()
    markers = [".workbuddy/skills/leadership-report-generator",
               ".workbuddy/skills/knowledge-base-organizer"]
    for p in [cur, *cur.parents]:
        if any((p / m).is_dir() for m in markers):
            return p
    raise SystemExit(
        "[error] 无法定位工作区。请用 --workspace 指定，"
        "或设置环境变量 WATER_WORKSPACE。"
    )


def plugin_roots(workspace: Path) -> dict[str, Path]:
    """插件层两副本（P ↔ R）。Q 无插件外壳，不参与。"""
    return {
        "P": Path.home() / ".workbuddy/plugins/marketplaces/my-experts/plugins/"
                            "water-office-writer",
        "R": workspace / "专家分发/company-marketplace/plugins/water-office-writer",
    }


# ---------------------------------------------------------------- 工具

def digest(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()[:12] if p.is_file() else "-"


def ignored(rel: str, cfg: SkillCfg | None) -> bool:
    if cfg is not None:
        if rel in cfg.ignore_relnames:
            return True
        if rel.startswith(cfg.ignore_relprefix):
            return True
    if Path(rel).suffix.lower() in IGNORE_SUFFIX:
        return True
    return any(part in IGNORE_DIRS for part in rel.split("/"))


def walk(root: Path, cfg: SkillCfg) -> list[str]:
    out: list[str] = []
    if not root.is_dir():
        return out
    for dp, dns, fns in os.walk(root):
        dns[:] = [d for d in dns if d not in IGNORE_DIRS]
        for fn in fns:
            rel = str((Path(dp) / fn).relative_to(root)).replace("\\", "/")
            if not ignored(rel, cfg):
                out.append(rel)
    return sorted(out)


def unfold_cr(data: bytes) -> tuple[bytes, int]:
    """把历史腐坏的双 CR 行尾 `\\r\\r\\n` 折回 `\\r\\n`。

    返回 (新字节, 折叠处数)。幂等：`\\r\\r\\n` 在正常 UTF-8 文本里不该出现，
    折叠是确定且无损的。

    为什么必须在**读取端**做：腐坏是"多副本同时中招"（同一份内容被同一段
    有缺陷的代码写过），所以从任一副本派生的期望值都会继承腐坏 —— 直接照抄
    等于把破坏洗白成"一致"。读取端折叠才能让 `--apply` 真正自愈。
    """
    n = data.count(b"\r\r\n")
    if not n:
        return data, 0
    return data.replace(b"\r\r\n", b"\r\n"), n


def transform(text: str, dst: str, cfg: SkillCfg, rel: str = "") -> tuple[str, str]:
    """把 P 出身的文本改写为目标副本 `dst` 的身份。

    返回 (改写后文本, 告警)。`cfg.rename=False` 的技能三副本同名，原样返回。
    """
    if not cfg.rename or rel in cfg.no_transform:
        return text, ""
    has_p, has_q = NAME_P in text, NAME_Q in text
    if has_p and has_q:
        return text, f"{rel}: 文本内两个技能名并存，已跳过身份替换"
    if dst == "Q":
        return text.replace(TITLE_P, TITLE_Q).replace(NAME_P, NAME_Q), ""
    return text.replace(TITLE_Q, TITLE_P).replace(NAME_Q, NAME_P), ""


def expected_text(rel: str, src_root: Path, dst_root: Path, dst_code: str,
                  cfg: SkillCfg) -> tuple[str, str]:
    """给定源文件，算出副本 `dst_code` 应具有的文本内容。

    与 `transform` 的区别：`cfg.path_rewrite_rel` 指定的文件内含绝对路径，
    各副本天生不同，因此按"路径已重写为本副本实际位置"的形态产出期望值，
    避免误报漂移。

    ⚠️ 用 `read_bytes().decode("utf-8")` 而非 `read_text()`：后者会做**通用换行
    归一化**（把 `\\r\\n` 变成 `\\n`），抹掉源文件的行尾风格，导致写回时行尾漂移。
    """
    raw = src_root.joinpath(rel).read_bytes()
    txt, warn = transform(unfold_cr(raw)[0].decode("utf-8"), dst_code, cfg, rel)
    if cfg.path_rewrite_rel is None or rel != cfg.path_rewrite_rel:
        return txt, warn
    # json.dumps 只产出 LF。按源文件实际行尾风格还原，并保留其尾换行有无。
    crlf = "\r\n" in txt
    data = json.loads(txt)
    cdir = dst_root / "selftest/cases"
    for item in data:
        cand = sorted(cdir.glob(f"{item.get('case_id', '')}_*.md")) if cdir.is_dir() else []
        if cand:
            item["md_path"] = str(cand[0])
    out = json.dumps(data, ensure_ascii=False, indent=2)
    if crlf:
        out = out.replace("\n", "\r\n")
    if txt.endswith(("\n", "\r")):
        out += "\r\n" if crlf else "\n"
    return out, warn


def expected_bytes(rel: str, src_root: Path, dst_root: Path, dst_code: str,
                   cfg: SkillCfg) -> bytes:
    """副本 `dst_code` 应有的**字节**内容。

    ⚠️ 全链路走 `read_bytes` / `write_bytes`，绝不经过 `Path.read_text` / `write_text`：
    后者在 Windows 上会把 `\\n` 翻译成 `\\r\\n`，若原文已是 CRLF 就会写出 `\\r\\r\\n`
    —— `str.splitlines()` 会把这种行尾当成**两个**换行，导致 Markdown 表格行之间
    出现空行、表格落版失效（2026-09-14 实际踩坑，见 SKILL.md §十）。
    """
    return expected_text(rel, src_root, dst_root, dst_code, cfg)[0].encode("utf-8")


# ---------------------------------------------------------------- 校验

def verify_copy(code: str, root: Path, cfg: SkillCfg | None = None) -> list[str]:
    """同步后语法校验：YAML / JSON 可解析，Python 可编译，行尾无异常。"""
    problems: list[str] = []
    for rel in walk(root, cfg):
        p = root / rel
        ext = p.suffix.lower()
        try:
            data = p.read_bytes()
            if b"\r\r\n" in data:
                problems.append(
                    f"[{code}] {rel}: 检出 \\r\\r\\n 异常行尾（会破坏 Markdown 表格解析）")
            text = data.decode("utf-8") if ext in TEXT_EXT else None
            if text is None:
                continue
            if ext in {".yaml", ".yml"}:
                import yaml
                yaml.safe_load(text)
            elif ext == ".json":
                json.loads(text)
            elif ext == ".py":
                ast.parse(text, filename=str(p))
        except Exception as e:  # noqa: BLE001
            problems.append(f"[{code}] {rel}: {type(e).__name__}: {e}")
    return problems


# ---------------------------------------------------------------- skill 层流程

def collect_actions(cfg: SkillCfg, only: list[str] | None):
    """巡检：返回 (需要同步的动作 [(rel, src_code)], 告警/说明列表)。"""
    actions: list[tuple[str, str]] = []
    notes: list[str] = []
    rs = cfg.roots

    allrels = sorted(set(walk(rs["P"], cfg)) | set(walk(rs["Q"], cfg)) | set(walk(rs["R"], cfg)))
    for rel in allrels:
        if only and rel.split("/")[0] not in only:
            continue
        src = cfg.take_from.get(rel, "P")
        sp = rs[src] / rel
        if not sp.is_file():
            notes.append(f"{rel}: 台账声明的源 {src} 中不存在 —— 请检查 TAKE_FROM")
            continue

        is_text = sp.suffix.lower() in TEXT_EXT
        for dst in ("P", "Q", "R"):
            dp = rs[dst] / rel
            if not is_text:                       # 二进制：直接比哈希
                if digest(dp) != digest(sp):
                    actions.append((rel, src))
                    break
            else:                                 # 文本：按目标身份比内容（字节级）
                expect, warn = expected_text(rel, rs[src], rs[dst], dst, cfg)
                if warn and warn not in notes:
                    notes.append(warn)
                if not dp.is_file() or dp.read_bytes() != expect.encode("utf-8"):
                    actions.append((rel, src))
                    break
    return actions, notes


def do_sync(cfg: SkillCfg, actions) -> list[str]:
    """把每个动作项的源内容写入三副本。

    ⚠️ 一律用 `write_bytes`，绝不 `write_text` —— 见 `expected_bytes` 的说明。
    """
    done: list[str] = []
    rs = cfg.roots
    for rel, src in actions:
        sp = rs[src] / rel
        is_text = sp.suffix.lower() in TEXT_EXT
        for dst in ("P", "Q", "R"):
            dp = rs[dst] / rel
            dp.parent.mkdir(parents=True, exist_ok=True)
            if not is_text:
                shutil.copy2(sp, dp)
            else:
                dp.write_bytes(expected_bytes(rel, rs[src], rs[dst], dst, cfg))
        done.append(f"{rel}  <- {src}")
    return done


def scan_corrupt(specs: list[SkillCfg]) -> list[str]:
    """各副本逐文件扫 `\\r\\r\\n`，返回问题清单（只读）。"""
    hits: list[str] = []
    for cfg in specs:
        for code, root in cfg.roots.items():
            if not root.is_dir():
                continue
            for rel in walk(root, cfg):
                p = root / rel
                if p.suffix.lower() not in TEXT_EXT:
                    continue
                try:
                    _, n = unfold_cr(p.read_bytes())
                except OSError:
                    continue
                if n:
                    hits.append(f"[{cfg.key}/{code}] {rel}: {n} 处 \\r\\r\\n")
    return hits


def prune(specs: list[SkillCfg]) -> list[str]:
    """删除 *.bak / __pycache__（各技能各副本）。"""
    removed = []
    seen: set[Path] = set()
    for cfg in specs:
        for root in cfg.roots.values():
            if not root.is_dir() or root in seen:
                continue
            seen.add(root)
            for p in list(root.rglob("*.bak")) + list(root.rglob("__pycache__")):
                if p.is_dir():
                    shutil.rmtree(p, ignore_errors=True)
                else:
                    p.unlink(missing_ok=True)
                removed.append(f"{p}")
    return removed


# ---------------------------------------------------------------- 插件层流程

# 插件层台账：{相对路径: 取自哪个副本}。空 = 全部以 P 为准。
#
# 2026-09-14 裁决依据（各项均已实测，非推测）：
#   agents/water-office-writer.md  P 为升级版（语言手册 / 三文种分道 / 金句密度红线 /
#                                  语料短板声明），R 为 v1.5 旧版
#   .codebuddy-plugin/plugin.json  P 已是 v1.6 元数据，R 仍标 1.5.0 且描述含
#                                  v1.2–v1.5 过时"新增"标记
#   USER_GUIDE.md / README.md      职称统一为「总裁」（2026-09-14 用户确认
#                                  「执行总裁」为误称，统一改为「总裁」）
#   avatars/expert.png              P（204KB）为集团数字员工形象图的 512×512 重制版，
#                                  与官方图 8×8 感知哈希仅差 5/64；R（399KB）是
#                                  初始占位图（与官方图差 21/64）
PLUGIN_TAKE_FROM: dict[str, str] = {}

# 已知存在差异但**尚未裁决**的文件：只报告，不自动同步（避免把猜测写成事实）。
PLUGIN_UNRESOLVED: set[str] = set()

PLUGIN_IGNORE_DIRS = {"skills", "output", "__pycache__", ".git", "node_modules"}


def collect_plugin_actions(prs: dict[str, Path], notes: list[str]) -> list[tuple[str, str]]:
    """插件层巡检（纯字节比对，无身份替换）。"""
    actions: list[tuple[str, str]] = []

    def _walk(root: Path) -> list[str]:
        out: list[str] = []
        if not root.is_dir():
            return out
        for dp, dns, fns in os.walk(root):
            dns[:] = [d for d in dns if d not in PLUGIN_IGNORE_DIRS]
            for fn in fns:
                rel = str((Path(dp) / fn).relative_to(root)).replace("\\", "/")
                if Path(rel).suffix.lower() not in IGNORE_SUFFIX:
                    out.append(rel)
        return sorted(out)

    allrels = sorted(set(_walk(prs["P"])) | set(_walk(prs["R"])))
    for rel in allrels:
        if rel in PLUGIN_UNRESOLVED:
            continue
        src = PLUGIN_TAKE_FROM.get(rel, "P")
        sp = prs[src] / rel
        if not sp.is_file():
            notes.append(f"[plugin] {rel}: 台账声明的源 {src} 中不存在 —— 请检查 PLUGIN_TAKE_FROM")
            continue
        want = unfold_cr(sp.read_bytes())[0]
        for dst in ("P", "R"):
            dp = prs[dst] / rel
            if not dp.is_file() or dp.read_bytes() != want:
                actions.append((rel, src))
                break
    return actions


def do_sync_plugin(prs: dict[str, Path], actions) -> list[str]:
    done: list[str] = []
    for rel, src in actions:
        data = unfold_cr((prs[src] / rel).read_bytes())[0]
        for dst in ("P", "R"):
            dp = prs[dst] / rel
            dp.parent.mkdir(parents=True, exist_ok=True)
            dp.write_bytes(data)
        done.append(f"{rel}  <- {src}")
    return done


def unresolved_report(prs: dict[str, Path]) -> list[str]:
    """列出 `PLUGIN_UNRESOLVED` 中当前确实仍有差异的项。"""
    out: list[str] = []
    for rel in sorted(PLUGIN_UNRESOLVED):
        a, b = digest(prs["P"] / rel), digest(prs["R"] / rel)
        if a != b:
            out.append(f"{rel}: P={a} R={b}")
    return out


# ---------------------------------------------------------------- 主流程

def main() -> int:
    ap = argparse.ArgumentParser(description="文秘专员 跨副本一致性巡检与同步")
    ap.add_argument("--workspace", help="工作区根（默认自动向上查找）")
    ap.add_argument("--apply", action="store_true", help="执行同步（默认仅巡检）")
    ap.add_argument("--check", action="store_true", help="只巡检（默认行为）")
    ap.add_argument("--only", nargs="*", help="只处理这些顶层目录，如 scripts assets")
    ap.add_argument("--skill", nargs="*", help="只处理这些技能（默认全部注册技能）")
    ap.add_argument("--no-verify", action="store_true", help="跳过同步后语法校验")
    ap.add_argument("--no-plugin", action="store_true", help="不巡检插件层（skill 层之上）")
    ap.add_argument("--prune", action="store_true", help="附带删除 *.bak / __pycache__")
    args = ap.parse_args()

    ws = Path(args.workspace).resolve() if args.workspace else find_workspace()
    prs = plugin_roots(ws)

    specs = [c for c in SKILLS if not args.skill or c.key in args.skill]
    unknown = set(args.skill or []) - {c.key for c in SKILLS}
    if unknown:
        print(f"[error] 未注册的技能：{sorted(unknown)}；已注册：{[c.key for c in SKILLS]}",
              file=sys.stderr)
        return 2
    for c in specs:
        c.resolve_roots(ws)

    missing = [(c.key, k) for c in specs for k, r in c.roots.items() if not r.is_dir()]
    if missing:
        print(f"[error] 副本目录不存在：{missing}", file=sys.stderr)
        return 2

    print(f"工作区：{ws}")
    for c in specs:
        print(f"\n[{c.key}] {c.label}")
        for k in ("P", "Q", "R"):
            print(f"  {k} = {c.roots[k]}")
    if not args.no_plugin:
        for k in ("P", "R"):
            if prs[k].is_dir():
                print(f"\n[plugin] {k} = {prs[k]}")

    if args.prune:
        rm = prune(specs)
        print(f"\n--- 清理 {len(rm)} 项 ---")
        for x in rm:
            print("  -", x)

    rc = 0
    plan: list[tuple[SkillCfg, list[tuple[str, str]]]] = []

    # ------------------------------------------------ skill 层（逐技能）
    for cfg in specs:
        print(f"\n=== skill 层 · {cfg.key}（P / Q / R 三副本）===")
        actions, notes = collect_actions(cfg, args.only)
        for n in notes:
            print("[warn]", n)
        if actions:
            print(f"--- 漂移 {len(actions)} 项 ---")
            for rel, src in actions:
                print(f"  {rel}  <- {src}")
                rs = cfg.roots
                print(f"      P={digest(rs['P']/rel)}  Q={digest(rs['Q']/rel)}  R={digest(rs['R']/rel)}")
        else:
            print("[OK] 三副本一致，无漂移。")
        plan.append((cfg, actions))

    corrupt = scan_corrupt(specs)
    if corrupt:
        print(f"\n--- 行尾腐坏 {len(corrupt)} 处（\\r\\r\\n，会破坏 Markdown 表格）---")
        for x in corrupt:
            print("  !", x)

    # ------------------------------------------------ 插件层（P ↔ R）
    plug_actions: list[tuple[str, str]] = []
    if not args.no_plugin:
        print("\n=== 插件层（P ↔ R 两副本）===")
        pnotes: list[str] = []
        plug_actions = collect_plugin_actions(prs, pnotes)
        for n in pnotes:
            print("[warn]", n)
        for x in unresolved_report(prs):
            print(f"[note] 待裁决，未自动同步：{x}")
        if plug_actions:
            print(f"--- 漂移 {len(plug_actions)} 项 ---")
            for rel, src in plug_actions:
                print(f"  {rel}  <- {src}")
                print(f"      P={digest(prs['P']/rel)}  R={digest(prs['R']/rel)}")
        else:
            print("[OK] 插件层一致，无漂移。")

    drifted = any(a for _, a in plan) or bool(plug_actions) or bool(corrupt)
    if not args.apply:
        if drifted:
            rc = 1
            print("\n[check] 仅巡检。加 --apply 执行同步。")
        else:
            print("\n[check] 全部一致，无需同步。")
        return rc

    if not drifted:
        print("\n[OK] 无需同步。")
        return 0

    # ------------------------------------------------ 执行同步
    for cfg, actions in plan:
        if actions:
            done = do_sync(cfg, actions)
            print(f"\n--- [{cfg.key}] 已同步 {len(done)} 个文件 × 3 副本 ---")
    if plug_actions:
        pdone = do_sync_plugin(prs, plug_actions)
        print(f"--- 插件层已同步 {len(pdone)} 个文件 × 2 副本 ---")

    if not args.no_verify:
        problems: list[str] = []
        for cfg in specs:
            for code, root in cfg.roots.items():
                problems += verify_copy(code, root, cfg)
        if not args.no_plugin:
            for code, root in prs.items():
                problems += verify_copy(code, root, None)
        if problems:
            print(f"\n[FAIL] 校验发现 {len(problems)} 个问题：")
            for x in problems:
                print("  x", x)
            return 1
        print("\n[OK] 校验通过（YAML / JSON 可解析，Python 可编译）。")

    # ------------------------------------------------ 复检
    for cfg, _ in plan:
        rest, _ = collect_actions(cfg, args.only)
        if rest:
            print(f"\n[warn] [{cfg.key}] 同步后仍有 {len(rest)} 项漂移：")
            for rel, src in rest:
                print(f"  {rel}  <- {src}")
            return 1
    left = scan_corrupt(specs)
    if left:
        print(f"\n[FAIL] 同步后仍残留 {len(left)} 处行尾腐坏：")
        for x in left:
            print("  x", x)
        return 1
    if not args.no_plugin:
        rest2 = collect_plugin_actions(prs, [])
        if rest2:
            print(f"\n[warn] 同步后仍有 {len(rest2)} 项插件层漂移：")
            for rel, src in rest2:
                print(f"  {rel}  <- {src}")
            return 1
    print("[OK] 同步后复检：全部技能三副本 + 插件层两副本均一致，行尾无腐坏。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
