#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""kbsync.py —— 乐享知识库增量同步引擎（L3）

职责单一：把**本地已归好类的知识库**镜像到乐享的某个知识库空间里。
分类判定不在这里做 —— 那是 intake.py / classify.py 的事。本脚本只回答一个问题：

    「这些本地文件，哪些还没推上去？推的时候要落在哪个远端目录？」

================================================================================
设计要点
================================================================================

【1】授权域（authorized domain）是一切的前提
    没有 `_authorized_domains.json` 里的授权记录，脚本**拒绝执行同步**。
    用户把根 `space_id` 交出来的那一刻，即已把"该空间内的目录路由"委托给我们：
    域内子目录自主路由属【委托决策】，不是官方安全规则所禁止的"自行选择目标"。
    域外（新主空间 / 新知识库）一律走 --authorize 确认流程。

【2】对齐乐享官方写入红线（不可绕过）
    · 禁止遍历团队/知识库列表后自行选目标 → `--discover` 只读展示，由人决定
    · 禁止按名称"看起来合适"就写入         → 目标由本地目录结构镜像而来，不猜
    · 禁止未确认就写入                     → 默认 dry-run，`--apply` 才落笔
    · 只创建，不删除、不覆盖                → 远端同名条目一律复用，不覆盖
    · 凭据不落盘                           → token 由 lexiang_client 从连接器自动发现

【3】幂等：以内容 md5 为主键
    读 `_sync_state.json`，记录里已有 `remote_entry_id` 就跳过。
    所以"重跑一次"永远是安全的 —— 第二次跑全部输出"跳过（幂等）"。
    若台账被删而远端已有同名条目，会用**同名复用**自愈，而不会传出一堆
    `xxx(1).docx`。这条兜底比"信任台账"更稳。

【4】分批：官方规范 `文件/条目数量 > 20 时每批 ≤ 20`
    每批结束落一次盘（状态 + 目录缓存 + 审计日志），中途中断也不会丢进度。

【5】三本账各记各的
    _intake_log.csv  入库决策  |  _sync_state.json  幂等状态  |  _sync_log.csv  传输事实

================================================================================
用法
================================================================================

  # 0. 自检连通性 / 取账号与域名信息
  python lexiang_client.py

  # 1. 首次授权（把乐享知识库的 space_id 交给它，只需做一次）
  python kbsync.py --kb-root <知识库根> --space-id <space_id> --authorize

  # 2. 看计划（默认 dry-run，不会写任何东西）
  python kbsync.py --kb-root <知识库根>

  # 3. 真正同步
  python kbsync.py --kb-root <知识库根> --apply

  # 4. 只推 5 份试水 / 只推某个子树 / 强制重传
  python kbsync.py --kb-root <知识库根> --apply --limit 5
  python kbsync.py --kb-root <知识库根> --apply --subpath "01-龚利民（董事长）/集团半年工作总结"
  python kbsync.py --kb-root <知识库根> --apply --force

退出码：0 成功 · 2 参数/鉴权错误 · 3 未授权 · 4 部分失败
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
import time
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path, PurePosixPath
from typing import Any, Iterable

# 允许作为脚本直接运行（scripts/ 下的兄弟模块互相 import）
_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

import lexiang_client as LC  # noqa: E402
import state as S  # noqa: E402

if sys.stdout.encoding is None or sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

# 允许上传的扩展名。
# 刻意**与 classify.SUPPORTED_EXT 分开定义**：分类引擎关心"能不能解析出文字"，
# 同步引擎关心"能不能传上去"，两者职责不同。这里额外允许图片与表格。
SYNC_EXT = {".docx", ".doc", ".pptx", ".ppt", ".pdf", ".md", ".txt",
            ".xlsx", ".xls", ".png", ".jpg", ".jpeg", ".gif", ".zip"}

DEFAULT_BATCH_SIZE = 20          # 官方规范上限
MAX_TAGS = 6

# taxonomy.yaml 读不到时的兜底词表（内容与 taxonomy 的 tag_vocabulary 一致）
DEFAULT_VOCAB = {
    "主题": ["鸿蒙", "AI", "人工智能", "智慧水务", "国际化", "出海", "新质生产力", "十五五"],
    "品牌": ["千家万户水管家", "深水云脑", "深小水AI智能体", "水鸿生态"],
    "机构": ["住建部", "水利部", "生态环境部", "中国水协", "供排水协会", "粤港澳大湾区", "GIIC"],
}

# taxonomy.yaml 读不到时的兜底别名（与 taxonomy: spaces[].aliases 一致）
DEFAULT_ALIASES = {
    "01-龚利民（董事长）": ["01-龚利民", "龚利民-董事长", "龚利民"],
    "02-方琳（总裁）": ["02-方琳（执行总裁）", "02-方琳", "方琳-总裁", "方琳"],
}


# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------


def _first(d: Any, keys: Iterable[str], default: Any = None) -> Any:
    """按顺序取第一个非空字段 —— 乐享各接口的字段命名不完全统一。"""
    if not isinstance(d, dict):
        return default
    for k in keys:
        v = d.get(k)
        if v not in (None, "", [], {}):
            return v
    return default


def entry_id_of(entry: Any) -> str:
    """从各种返回结构里挖出条目 id。"""
    if not isinstance(entry, dict):
        return ""
    for k in ("id", "entry_id", "entryId", "_id"):
        v = entry.get(k)
        if isinstance(v, str) and v:
            return v
    for wrapper in ("entry", "space", "item", "data", "result"):
        v = entry.get(wrapper)
        if isinstance(v, dict):
            got = entry_id_of(v)
            if got:
                return got
    return ""


def _is_folder(ch: Any) -> bool:
    if not isinstance(ch, dict):
        return False
    t = str(_first(ch, ("entry_type", "type", "kind"), "")).lower()
    if t:
        return t in ("folder", "dir", "directory")
    # 没有类型字段时按"有没有文件特征"反推
    return not _first(ch, ("file_id", "fileId", "size", "mime_type", "mimeType"))


def _extract_upload(resp: Any) -> tuple[str, str]:
    """从 file_apply_upload 的返回里取出 (upload_url, session_id)。"""
    if not isinstance(resp, dict):
        return "", ""
    url = _first(resp, ("upload_url", "uploadUrl", "presigned_url", "presignedUrl", "url")) or ""
    sid = _first(resp, ("session_id", "sessionId", "id")) or ""
    for w in ("session", "upload", "file", "data"):
        v = resp.get(w)
        if isinstance(v, dict):
            url = url or (_first(v, ("upload_url", "uploadUrl", "presigned_url", "presignedUrl", "url")) or "")
            sid = sid or (_first(v, ("session_id", "sessionId", "id")) or "")
    return str(url), str(sid)


def space_info(raw: Any) -> tuple[str, str]:
    """从 space_describe_space 返回里取出 (空间名, root_entry_id)。"""
    sp = raw.get("space") if isinstance(raw, dict) and isinstance(raw.get("space"), dict) else raw
    if not isinstance(sp, dict):
        return "-", ""
    name = str(_first(sp, ("name", "title", "space_name"), "-") or "-")
    root = str(_first(sp, ("root_entry_id", "rootEntryId", "root_id"), "") or "")
    if not root:
        for k in ("root_entry", "root"):
            v = sp.get(k)
            if isinstance(v, dict):
                root = entry_id_of(v)
            elif isinstance(v, str):
                root = v
            if root:
                break
    return name, root


def build_url(domain: str, company_from: str, entry_id: str) -> str:
    """按官方 URL 规则拼接条目访问链接。

    规则（references/base.md）：
      · 域名含三级前缀（如 csig.lexiangla.com）→ {domain}/pages/{id}
      · 域名为顶级域名（如 lexiangla.com）      → {domain}/pages/{id}?company_from=xxx
    ⛔ 绝不能用 MCP endpoint 域名拼接用户访问链接。
    """
    if not domain or not entry_id:
        return ""
    d = domain.rstrip("/")
    host = d.split("://", 1)[-1]
    if host.count(".") >= 2:
        return f"{d}/pages/{entry_id}"
    return f"{d}/pages/{entry_id}?company_from={company_from}" if company_from else f"{d}/pages/{entry_id}"


def fmt_size(n: int) -> str:
    f = float(n)
    for unit in ("B", "KB", "MB", "GB"):
        if f < 1024 or unit == "GB":
            return f"{f:.1f} {unit}" if unit != "B" else f"{int(f)} B"
        f /= 1024
    return f"{f:.1f} GB"


# ---------------------------------------------------------------------------
# 文档同源判定
# ---------------------------------------------------------------------------
#
# 【为什么需要这一层】
# 远端 `数字员工_知识库` 是人工先建好的，文件名与本地**并不逐一对应**。实测到的差异：
#   · 扩展名不进 name 字段（`xxx.docx` 的 name 是 `xxx`，扩展名在 extension 里）
#   · 日期被人工修正过：本地 20260701 → 远端 20260801
#   · 标题被缩写：本地「第六届粤港澳大湾区水务论坛暨第十五届深港珠澳供水届学术交流会致辞」
#                → 远端「第六届粤港澳大湾区水务论坛致辞」
#   · 格式被转换：本地 .pptx → 远端 .pdf
# 如果只按名字全等判断"是否已存在"，首次同步会把十几份文件**重复传一遍**，
# 制造一堆 `xxx(1)`。所以这里用「文档身份」而不是「文件名」来判同源。
#
# 【判定规则（按可信度分档）】
#   tier1 exact   文件名完全一致
#   tier2 strong  同类型 + 同主体 + 标题相似度 ≥ 0.85      → 认定为同一份
#   tier3 weak    同类型 + 同主体 + 同日期 + 标题相似度 ≥ 0.50 → 疑似同一份，**交人确认**
# 类型/主体解析不出来时，退化用整体名称相似度 ≥ 0.92 兜底。
#
# 【一对一贪心】同一目录内，一份远端条目只能被一份本地文件认领。
# 否则「同一天两篇致辞」这种会互相抢。

TITLE_SIM_STRONG = 0.85
TITLE_SIM_WEAK = 0.50
NAME_SIM_FALLBACK = 0.92

_STRIP_RE = re.compile(
    r"[\s_\-—－–~～·、,，.。;；:：!！?？“”\"'‘’()（）\[\]【】{}<>《》/\\|+]+")

_DATE_RE = re.compile(r"^\d{8}$")


def norm_text(s: Any) -> str:
    """归一化：去掉空白与各类分隔符、统一小写 —— 用于名称相似度比较。"""
    return _STRIP_RE.sub("", str(s or "")).lower()


def title_sim(a: Any, b: Any) -> float:
    na, nb = norm_text(a), norm_text(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    return SequenceMatcher(None, na, nb).ratio()


def split_docname(stem: str) -> tuple[str, str, str, str]:
    """按命名模板 `{date}_{type}_{title}_{subject}` 拆解文件名主干。

    返回 (date, type, title, subject)；拆不出就退化为 ("", "", stem, "")。
    本地与远端用**同一套拆解规则**，所以两边的 title/subject 位置是一致的。
    """
    parts = [p for p in str(stem or "").split("_")]
    if len(parts) >= 4:
        return parts[0], parts[1], "_".join(parts[2:-1]), parts[-1]
    if len(parts) == 3:
        return parts[0], parts[1], parts[2], ""
    return "", "", str(stem or ""), ""


def doc_match(local_stem: str, remote_name: str) -> tuple[float, str, str]:
    """判断本地文件主干与远端条目名是否同源。返回 (分数, 档位, 说明)。"""
    if norm_text(local_stem) and norm_text(local_stem) == norm_text(remote_name):
        return 1.0, "exact", "文件名完全一致"

    ld, lt, ltitle, lsub = split_docname(local_stem)
    rd, rt, rtitle, rsub = split_docname(remote_name)
    same_kind = (norm_text(lt) and norm_text(lt) == norm_text(rt))
    same_subj = (norm_text(lsub) and norm_text(lsub) == norm_text(rsub))

    if same_kind and same_subj:
        sim = title_sim(ltitle, rtitle)
        if sim >= TITLE_SIM_STRONG:
            return 0.98, "strong", f"同类型同主体，标题高度一致（{sim:.2f}）"
        if ld and rd and ld == rd and sim >= TITLE_SIM_WEAK:
            return 0.80, "weak", f"同类型同主体同日期，标题相近（{sim:.2f}）"

    sim = title_sim(local_stem, remote_name)
    if sim >= NAME_SIM_FALLBACK:
        return 0.90, "strong", f"整体名称高度相似（{sim:.2f}）"
    return 0.0, "", ""


def find_same_doc(index: "RemoteIndex", parent_id: str, local_stem: str,
                  used: set[str]) -> tuple[dict | None, str, str, str]:
    """在某个远端目录里给本地文件找同源条目。

    返回 (远端条目, entry_id, 档位, 说明)；找不到返回 (None, "", "", "")。
    `used` 记录已被认领的远端 entry_id，保证一对一。
    """
    best: tuple[float, dict, str, str] | None = None
    for ch in index.children(parent_id):
        if _is_folder(ch):
            continue
        eid = entry_id_of(ch)
        if not eid or eid in used:
            continue
        sc, tier, reason = doc_match(local_stem, str(_first(ch, ("name", "title"), "")))
        if sc > 0 and (best is None or sc > best[0]):
            best = (sc, ch, tier, reason)
    if best is None:
        return None, "", "", ""
    return best[1], entry_id_of(best[1]), best[2], best[3]


# ---------------------------------------------------------------------------
# 授权域台账 _authorized_domains.json
# ---------------------------------------------------------------------------


class AuthStore:
    """授权域记录。

    结构与 `routing-rules.yaml: authorized_domain` 的论证一一对应：
    一个授权域 = 一个 space_id + 它的 root_entry_id + 域内已建目录缓存。

    目录缓存（folders）放在这里而不是 sync_state 里，理由：
      它描述的是**远端空间的结构**，不是本地文件的状态。
      换一个授权域，这张缓存就该换一份，二者生命周期一致。
    """

    VERSION = 1

    def __init__(self, kb_root: Path):
        self.root = Path(kb_root)
        self.path = self.root / S.AUTH_NAME
        self.data: dict[str, Any] = {"version": self.VERSION, "updated": "", "domains": []}
        self.load()

    def load(self) -> None:
        if not self.path.exists():
            return
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
            if isinstance(raw.get("domains"), list):
                self.data = raw
                self.data.setdefault("version", self.VERSION)
        except Exception as e:
            sys.stderr.write(f"[WARN] {self.path.name} 读取失败（按未授权处理）：{e}\n")

    def save(self) -> None:
        self.data["version"] = self.VERSION
        self.data["updated"] = S.now_iso()
        S.atomic_write_text(self.path, json.dumps(self.data, ensure_ascii=False, indent=2))

    @property
    def domains(self) -> list[dict]:
        return self.data.setdefault("domains", [])

    def get(self, space_id: str | None = None) -> dict | None:
        if not self.domains:
            return None
        if space_id:
            for d in self.domains:
                if d.get("space_id") == space_id:
                    return d
            return None
        return self.domains[0]

    def add(self, space_id: str, space_name: str, root_entry_id: str,
            operator: str = "user", notes: str = "",
            mirror_root_name: str = "", root_kind: str = "space") -> dict:
        dom = self.get(space_id)
        if dom is None:
            dom = {"space_id": space_id, "folders": {}}
            self.domains.append(dom)
        dom.update({
            "space_name": space_name or dom.get("space_name") or "-",
            "root_entry_id": root_entry_id,
            "mirror_root_name": mirror_root_name or space_name or "-",
            "root_kind": root_kind,     # space = 镜像到空间根；folder = 镜像到指定子目录
            "authorized_at": S.now_iso(),
            "authorized_by": operator,
            "scope": "域内子目录自主路由；域外一律确认",
        })
        if notes:
            dom["notes"] = notes
        dom.setdefault("folders", {})
        return dom

    def set_folders(self, space_id: str, folders: dict) -> None:
        dom = self.get(space_id)
        if dom is not None:
            dom["folders"] = folders
            dom["folders_at"] = S.now_iso()

    def drop(self, space_id: str) -> bool:
        before = len(self.domains)
        self.data["domains"] = [d for d in self.domains if d.get("space_id") != space_id]
        return len(self.data["domains"]) != before


# ---------------------------------------------------------------------------
# 标签
# ---------------------------------------------------------------------------


def load_taxonomy(skill_dir: Path | None) -> dict:
    """读 taxonomy.yaml —— 标签词表与目录别名都是资产，规则不硬编码。"""
    if not skill_dir:
        return {}
    p = Path(skill_dir) / "assets" / "taxonomy.yaml"
    if not p.exists():
        return {}
    try:
        import yaml  # 延迟导入：没有 PyYAML 也能跑同步（退化为兜底常量）
        return yaml.safe_load(p.read_text(encoding="utf-8")) or {}
    except Exception:
        return {}


def tag_vocab_of(doc: dict) -> dict:
    vocab = doc.get("tag_vocabulary") or {}
    out = {k: list(v) for k, v in vocab.items() if isinstance(v, list) and v}
    return out or dict(DEFAULT_VOCAB)


def space_aliases_of(doc: dict) -> dict[str, set[str]]:
    """本地主空间名 → 可接受的远端同义目录名。

    【为什么需要】远端目录是人工建的，名字未必和本地一字不差。
    实测到的例子：远端历史版本可能是 `02-方琳（执行总裁）`（已统一为 `02-方琳（总裁）`）。
    没有别名机制的话，同步会在远端**又建一个目录**，把结构搞乱 ——
    而这恰恰是本项目最想避免的事。别名登记在 taxonomy.yaml 里，属可积累资产。
    """
    out: dict[str, set[str]] = {}
    for sp in doc.get("spaces") or []:
        if not isinstance(sp, dict):
            continue
        sid = str(sp.get("id") or "").strip()
        if not sid:
            continue
        names = {sid}
        for a in sp.get("aliases") or []:
            if a:
                names.add(str(a).strip())
        out[sid] = names
    return out or {k: set(v) | {k} for k, v in DEFAULT_ALIASES.items()}


def build_tags(rec: dict, rel: str, vocab: dict) -> list[str]:
    """标签轴 = 年度 / 主体 / 类型 / 主题。

    「主题一律用标签，不建目录」（见 taxonomy.yaml 修改纪律），
    所以鸿蒙、AI、智慧水务这类词只在这里出现，不产生任何目录。
    """
    tags: list[str] = []

    def add(t: Any) -> None:
        t = str(t or "").strip()
        if t and t not in tags and t not in ("-", "无领导", "未判定", "None"):
            tags.append(t)

    year = str(rec.get("date") or "")[:4]
    if year.isdigit():
        add(f"{year}年")
    add(rec.get("subject"))
    add(rec.get("doc_type"))

    # 主题/品牌/机构：只在标题与文件名里找，避免从正文里误挂
    hay = " ".join([str(rec.get("title_clean") or ""), PurePosixPath(rel).name]).lower()
    for words in vocab.values():
        for w in words:
            if w and str(w).lower() in hay:
                add(w)
    return tags[:MAX_TAGS]


# ---------------------------------------------------------------------------
# 计划
# ---------------------------------------------------------------------------


@dataclass
class SyncItem:
    seq: int
    src: Path
    rel: str                       # 相对知识库根（posix，含文件名）
    md5: str
    size: int
    doc_type: str = ""
    subject: str = ""
    tags: list[str] = field(default_factory=list)
    status: str = "pending"        # pending / skipped / synced / reused / conflict / failed
    remote_entry_id: str = ""
    remote_url: str = ""
    remote_folder: str = ""        # 远端目录链（posix）
    match_tier: str = ""           # exact / strong / weak
    match_reason: str = ""
    folder_state: str = ""         # exists / create
    elapsed: float = 0.0
    notes: str = ""
    error: str = ""

    @property
    def stem(self) -> str:
        return PurePosixPath(self.rel).stem

    @property
    def parts(self) -> list[str]:
        """目标目录链（不含文件名）。"""
        p = PurePosixPath(self.rel)
        return list(p.parent.parts) if str(p.parent) != "." else []

    @property
    def remote_dir(self) -> str:
        return "/".join(self.parts)


def local_files(kb_root: Path, subpath: str | None = None) -> list[Path]:
    base = Path(kb_root)
    root = (base / subpath) if subpath else base
    if not root.exists():
        return []
    out: list[Path] = []
    for p in sorted(root.rglob("*")):
        if not p.is_file():
            continue
        try:
            rel_parts = p.relative_to(base).parts
        except ValueError:
            continue
        # `_` 开头的目录/文件 = 台账与内部产物，不进库
        if any(part.startswith("_") or part.startswith(".") for part in rel_parts):
            continue
        if p.suffix.lower() not in SYNC_EXT:
            continue
        out.append(p)
    return out


def build_plan(kb_root: Path, st: S.StateStore, vocab: dict,
               subpath: str | None = None, force: bool = False,
               limit: int | None = None) -> list[SyncItem]:
    items: list[SyncItem] = []
    for src in local_files(kb_root, subpath):
        rel = src.relative_to(kb_root).as_posix()
        md5 = S.md5_file(src)
        rec = st.get(md5) or {}
        size = src.stat().st_size
        it = SyncItem(
            seq=len(items) + 1, src=src, rel=rel, md5=md5, size=size,
            doc_type=str(rec.get("doc_type") or ""),
            subject=str(rec.get("subject") or ""),
            tags=build_tags(rec, rel, vocab),
        )
        if size == 0:
            it.status = "skipped"
            it.notes = "空文件，跳过"
        elif rec.get("remote_entry_id") and not force:
            # 台账说它已同步 —— 这是幂等的主要路径
            it.status = "skipped"
            it.remote_entry_id = str(rec["remote_entry_id"])
            it.remote_url = str(rec.get("remote_url") or "")
            it.notes = "已同步（幂等跳过）"
        items.append(it)

    if limit:
        pending = [i for i in items if i.status == "pending"][:limit]
        keep = {id(i) for i in pending}
        for i in items:
            if i.status == "pending" and id(i) not in keep:
                i.status = "skipped"
                i.notes = "超出 --limit，本轮不动"
    return items


# ---------------------------------------------------------------------------
# 远端索引（进程内缓存，减少 list 调用）
# ---------------------------------------------------------------------------


class RemoteIndex:
    """远端子条目缓存。

    一个 22 份文件的知识库如果每份都 list 一次父目录，就会有 22 次多余请求；
    按 parent 缓存后，同一目录下的 N 份文件只 list 一次。
    """

    def __init__(self, cli: LC.LexiangMCP):
        self.cli = cli
        self._cache: dict[str, list[dict]] = {}

    def children(self, parent_id: str) -> list[dict]:
        if parent_id not in self._cache:
            self._cache[parent_id] = self.cli.list_children(parent_id)
        return self._cache[parent_id]

    def invalidate(self, parent_id: str) -> None:
        self._cache.pop(parent_id, None)

    def find(self, parent_id: str, name: str, want_folder: bool | None = None) -> dict | None:
        for ch in self.children(parent_id):
            if str(_first(ch, ("name", "title"), "")) != name:
                continue
            if want_folder is None or _is_folder(ch) == want_folder:
                return ch
        return None

    def find_any(self, parent_id: str, names: Iterable[str],
                 want_folder: bool | None = None) -> tuple[dict | None, str]:
        """按一组同义名找子节点，返回 (节点, 命中的名字)。

        用于解决"远端目录名与本地不一致"（见 space_aliases_of 的说明）。
        """
        wanted = {str(n) for n in names if n}
        for ch in self.children(parent_id):
            nm = str(_first(ch, ("name", "title"), ""))
            if nm not in wanted:
                continue
            if want_folder is None or _is_folder(ch) == want_folder:
                return ch, nm
        return None, ""


def ensure_folder(index: RemoteIndex, cache: dict, root_id: str,
                  parts: list[str], aliases: dict | None = None) -> tuple[str, list[str], list[str]]:
    """逐级确保远端目录存在。

    返回 (末级目录 id, 本次新建的目录路径列表, 走别名复用的说明列表)。

    已经在文件夹缓存里的层级直接跳过 —— 所以第二次同步目录部分是零请求。
    匹配时同时接受 taxonomy 里登记的别名，避免在远端造出重复目录。
    """
    aliases = aliases or {}
    cur = root_id
    acc: list[str] = []
    created: list[str] = []
    aliased: list[str] = []
    for name in parts:
        acc.append(name)
        key = "/".join(acc)
        cached = cache.get(key)
        if cached:
            cur = cached
            continue
        found, hit = index.find_any(cur, aliases.get(name) or {name}, want_folder=True)
        if found:
            eid = entry_id_of(found)
            if hit and hit != name:
                aliased.append(f"{key} → 复用远端已有目录「{hit}」")
        else:
            ent = index.cli.create_entry(cur, name, "folder")
            eid = entry_id_of(ent)
            created.append(key)
        if not eid:
            raise LC.LexiangError(f"无法取得远端目录 [{key}] 的条目 id")
        cache[key] = eid
        cur = eid
    return cur, created, aliased


def resolve_folder_ro(index: RemoteIndex, cache: dict, root_id: str,
                      parts: list[str], aliases: dict | None = None) -> tuple[str, str]:
    """只读解析目录链 —— 任一层不存在就返回 ("", 第一个缺失的层级)。

    与 ensure_folder 的区别：**绝不创建**。用于 --inspect 预演，
    这样"看看会怎么同步"这个动作本身也是零写入的。
    """
    aliases = aliases or {}
    cur = root_id
    acc: list[str] = []
    for name in parts:
        acc.append(name)
        key = "/".join(acc)
        eid = cache.get(key)
        if not eid:
            found, _hit = index.find_any(cur, aliases.get(name) or {name}, want_folder=True)
            eid = entry_id_of(found) if found else ""
            if eid:
                cache[key] = eid
        if not eid:
            return "", key
        cur = eid
    return cur, ""


# ---------------------------------------------------------------------------
# 单份上传
# ---------------------------------------------------------------------------


def upload_one(cli: LC.LexiangMCP, item: SyncItem, parent_id: str,
               do_tags: bool, domain: str, company_from: str) -> None:
    """三步上传：申请 → PUT 原始字节 → 提交。"""
    mime = LC.mime_of(item.src)

    resp = cli.apply_upload(parent_id, item.src.name, item.size, mime)
    url, sid = _extract_upload(resp)
    if not url or not sid:
        raise LC.LexiangError(
            "file_apply_upload 返回不完整（缺 upload_url 或 session_id）："
            + json.dumps(resp, ensure_ascii=False)[:300])

    # put_blob 内部用流式 PUT —— 80MB+ 的 pptx 不能读进内存
    LC.LexiangMCP.put_blob(url, item.src, mime)

    ent = cli.commit_upload(sid)
    item.remote_entry_id = entry_id_of(ent)
    if not item.remote_entry_id:
        raise LC.LexiangError(
            "file_commit_upload 未返回条目 id："
            + json.dumps(ent, ensure_ascii=False)[:300])
    item.remote_url = str(_first(ent, ("url", "link", "entry_url", "web_url")) or "") \
        or build_url(domain, company_from, item.remote_entry_id)

    if do_tags and item.tags:
        try:
            cli.set_tags(item.remote_entry_id, item.tags)
        except Exception as e:  # 标签失败不应让整份文件算失败
            item.notes = (item.notes + f" · 标签写入失败：{e}").strip(" ·")


# ---------------------------------------------------------------------------
# 渲染
# ---------------------------------------------------------------------------


def print_plan(items: list[SyncItem], kb_root: Path, dom: dict,
               applied: bool, batch_size: int, force: bool,
               inspected: bool = False) -> None:
    counts: dict[str, int] = {}
    for it in items:
        counts[it.status] = counts.get(it.status, 0) + 1
    pend = counts.get("pending", 0)
    reuse = sum(1 for i in items if i.status == "pending" and i.match_tier in ("exact", "strong"))

    title = ("执行结果" if applied
             else ("预演（含远端比对，只读，不会写入任何内容）" if inspected
                   else "同步计划（dry-run，不会写入任何内容）"))
    print("=" * 84)
    print(f"乐享知识库增量同步 · {title}")
    print("=" * 84)
    print(f"  本地知识库   {kb_root}")
    print(f"  授权域       {dom.get('space_name')}  ({dom.get('space_id')})")
    print(f"  镜像根       {dom.get('mirror_root_name') or dom.get('space_name')}"
          + (f"   [{dom.get('root_entry_id')}]"))
    print(f"  共 {len(items)} 份：  " + " · ".join(f"{k} {v}" for k, v in sorted(counts.items())))
    if pend:
        batches = (pend + batch_size - 1) // batch_size
        line = f"  本轮待处理   {pend} 份，按每批 {batch_size} 份分 {batches} 批"
        if reuse:
            line += f"（其中 {reuse} 份判定为远端已存在，将复用不重传）"
        print(line)
        if force:
            print("  注意         --force 已开启：忽略台账与同源判定，一律重传")
    print()

    _ICON = {"pending": "[→] 待传", "skipped": "[=] 跳过", "synced": "[+] 已传",
             "reused": "[~] 复用", "conflict": "[!] 待确认", "failed": "[x] 失败"}
    shown = 0
    for it in items:
        if it.status == "skipped" and not it.remote_entry_id:
            continue
        print(f"  {_ICON.get(it.status, it.status)}  #{it.seq:<3} {PurePosixPath(it.rel).name}")
        meta = f"{it.doc_type or '未判定'} · {it.subject or '-'} · {fmt_size(it.size)}"
        if it.tags:
            meta += " · 标签 " + ",".join(it.tags)
        print(f"          {meta}")
        print(f"          → {it.remote_dir or '(根目录)'}"
              + ("   [目录不存在，将补建]" if it.folder_state == "create" else ""))
        if it.match_tier:
            print(f"          同源判定：{it.match_tier} · {it.match_reason}"
                  + (f" · 远端条目 {it.remote_entry_id}" if it.remote_entry_id else ""))
        if it.notes:
            print(f"          {it.notes}")
        if it.error:
            print(f"          [x] {it.error}")
        shown += 1
        if not applied and shown >= 40:
            rest = len(items) - shown
            if rest > 0:
                print(f"  …… 其余 {rest} 份省略")
            break
    print()


def print_conflicts(items: list[SyncItem]) -> None:
    cf = [i for i in items if i.status == "conflict"]
    if not cf:
        return
    print("=" * 84)
    print(f"需要你确认：{len(cf)} 份在远端疑似已有同源条目（命名/格式不一致）")
    print("=" * 84)
    for it in cf:
        print(f"  #{it.seq} 本地  {PurePosixPath(it.rel).name}")
        print(f"       远端  {it.remote_entry_id}  ← {it.match_reason}")
    print()
    print("  这三条路选一条：")
    print("    · 认定是同一份 → 加 --accept-conflicts（复用远端，不重传）")
    print("    · 确认是两份   → 加 --force（另传一份到同目录）")
    print("    · 先不动       → 什么都不加，这几份会一直挂在『待确认』")
    print()


def print_result(items: list[SyncItem], batch_id: str, elapsed: float) -> tuple[int, int, int]:
    ok = sum(1 for i in items if i.status in ("synced", "reused"))
    fail = sum(1 for i in items if i.status == "failed")
    skip = sum(1 for i in items if i.status == "skipped")
    cf = sum(1 for i in items if i.status == "conflict")
    total_bytes = sum(i.size for i in items if i.status == "synced")

    print("=" * 84)
    print(f"同步批次 {batch_id} 完成")
    print("=" * 84)
    print(f"  成功 {ok}  ·  跳过 {skip}  ·  失败 {fail}"
          + (f"  ·  待确认 {cf}" if cf else ""))
    if total_bytes:
        print(f"  本轮上传 {fmt_size(total_bytes)}，耗时 {elapsed:.1f}s"
              + (f"，均速 {fmt_size(int(total_bytes / elapsed))}/s" if elapsed > 0 else ""))
    for it in items:
        if it.status == "failed":
            print(f"  [x] {it.rel}\n      {it.error}")
    print()
    if fail:
        print("  失败不影响已成功部分：台账已逐批落盘，直接重跑即可续传。")
    elif not cf:
        print("  幂等成立：再跑一次将全部输出「跳过」。")
    print()
    return ok, skip, fail


# ---------------------------------------------------------------------------
# 子命令
# ---------------------------------------------------------------------------


def cmd_show_auth(kb_root: Path, as_json: bool) -> int:
    auth = AuthStore(kb_root)
    if as_json:
        print(json.dumps(auth.data, ensure_ascii=False, indent=2))
        return 0
    print("=" * 84)
    print("授权域记录")
    print("=" * 84)
    print(f"  文件  {auth.path}")
    if not auth.domains:
        print("  状态  未授权 —— 任何同步都会被拒绝")
        print()
        print("  授权方法：")
        print("    python kbsync.py --kb-root <知识库根> --space-id <space_id> --authorize")
        print("  space_id 取自乐享知识库链接：https://<域名>/spaces/<space_id>")
        return 0
    for d in auth.domains:
        folders = d.get("folders") or {}
        print(f"  空间       {d.get('space_name')}  ({d.get('space_id')})")
        print(f"  镜像根     {d.get('mirror_root_name') or d.get('space_name')}"
              + ("（指定子目录）" if d.get("root_kind") == "folder" else "（空间根）"))
        print(f"  根条目     {d.get('root_entry_id')}")
        print(f"  授权时间   {d.get('authorized_at')} · by {d.get('authorized_by')}")
        print(f"  目录缓存   {len(folders)} 个正式目录已映射"
              + (f"（{d.get('folders_at')} 刷新）" if d.get("folders_at") else ""))
        for k, v in folders.items():
            print(f"      {k}  →  {v}")
    return 0


def cmd_discover(limit: int = 20) -> int:
    """只读展示：把可见的团队与知识库列出来，供人自己挑。

    ⚠️ 这里刻意**不做任何选择、不做任何写入**。
    官方红线禁止"遍历列表后自行选目标"，本命令把选择权留给人：
    它只负责让你看见 id，看到之后由你显式执行 --authorize。
    """
    try:
        cli = LC.LexiangMCP()
        cli.handshake()
        me = cli.whoami()
    except LC.LexiangError as e:
        print(f"[FAIL] {e}")
        return 2

    print("=" * 84)
    print("只读探查（不会写入任何内容）")
    print("=" * 84)
    comp = me.get("company") or {}
    print(f"  账号   {me.get('name') or '-'}  <{me.get('email') or '-'}>")
    print(f"  公司   {comp.get('name') or '-'}   域名 {comp.get('company_domain') or '-'}")
    ps = me.get("personal_space_id")
    if ps:
        print(f"  个人知识库 space_id  {ps}")
    print()

    teams = []
    try:
        teams = cli.list_teams()
    except LC.LexiangError as e:
        print(f"  [WARN] 团队列表读取失败：{e}")
    if teams:
        print("  可选团队 / 知识库：")
        for t in teams[:limit]:
            tid = entry_id_of(t)
            print(f"    · {t.get('name') or '-'}   team_id={tid}")
            try:
                for sp in cli.list_spaces(tid)[:limit]:
                    print(f"        - {sp.get('name') or '-'}   space_id={sp.get('space_id') or entry_id_of(sp)}")
            except LC.LexiangError as e:
                print(f"        [WARN] {e}")
    print()
    print("  请把要同步到的 space_id 明确告知，然后执行：")
    print("    python kbsync.py --kb-root <知识库根> --space-id <space_id> --authorize")
    return 0


def cmd_authorize(kb_root: Path, space_id: str, operator: str,
                  notes: str, config: str | None,
                  root_entry_id: str = "") -> int:
    if not space_id:
        print("[FAIL] --authorize 必须同时给出 --space-id")
        print("       space_id 取自乐享知识库链接：https://<域名>/spaces/<space_id>")
        print("       不确定是哪个？先跑只读探查：python kbsync.py --discover")
        return 2
    if config:
        os.environ["KB_LEXIANG_CONFIG"] = config
    try:
        cli = LC.LexiangMCP()
    except LC.LexiangError as e:
        print(f"[FAIL] {e}")
        return 2

    try:
        cli.handshake()
        raw = cli.describe_space(space_id)
    except LC.LexiangError as e:
        print(f"[FAIL] 无法读取 space {space_id}：{e}")
        return 2

    name, space_root = space_info(raw)
    root, mirror_name, kind = space_root, name, "space"

    if root_entry_id:
        try:
            ent = cli.describe_entry(root_entry_id)
        except LC.LexiangError as e:
            print(f"[FAIL] 指定的 --root-entry-id 读取失败：{e}")
            return 2
        if entry_id_of(ent) != root_entry_id:
            print("[FAIL] --root-entry-id 不是有效的条目 id："
                  + json.dumps(ent, ensure_ascii=False)[:200])
            return 2
        belong = str(_first(ent, ("space_id",), "") or "")
        if belong and belong != space_id:
            print(f"[FAIL] 该条目属于 space {belong}，与你指定的 {space_id} 不一致，已停止。")
            return 2
        if not _is_folder(ent):
            print("[FAIL] --root-entry-id 必须指向一个目录（entry_type=folder）。")
            return 2
        root = root_entry_id
        mirror_name = str(_first(ent, ("name", "title"), "") or "-")
        kind = "folder"

    if not root:
        print("[FAIL] 该 space 未返回 root_entry_id，无法作为授权域根。")
        print("       原始返回：" + json.dumps(raw, ensure_ascii=False)[:300])
        return 2

    auth = AuthStore(kb_root)
    auth.add(space_id, name, root, operator=operator, notes=notes,
             mirror_root_name=mirror_name, root_kind=kind)
    auth.save()

    print("=" * 84)
    print("授权成功")
    print("=" * 84)
    print(f"  知识库空间   {name}")
    print(f"  space_id    {space_id}")
    print(f"  镜像根       {mirror_name}   （{'空间根' if kind == 'space' else '指定子目录'}）")
    print(f"  根条目       {root}")
    print(f"  记录位置     {auth.path}")
    print()
    print("  委托范围     域内子目录自主路由；域外（新空间/新知识库）一律需要重新授权")
    print("  下一步       python kbsync.py --kb-root \"%s\"" % kb_root)
    return 0


def cmd_tree(space_id: str, depth: int = 2, limit: int = 60,
             root_entry_id: str = "") -> int:
    """只读打印某个空间（或某个目录）下的目录树。

    和 --discover 一样，这是**给你挑目标用的**，不做任何选择、不做任何写入。
    """
    try:
        cli = LC.LexiangMCP()
        cli.handshake()
    except LC.LexiangError as e:
        print(f"[FAIL] {e}")
        return 2

    try:
        raw = cli.describe_space(space_id)
        name, root = space_info(raw)
    except LC.LexiangError as e:
        print(f"[FAIL] 无法读取 space {space_id}：{e}")
        return 2
    if root_entry_id:
        root = root_entry_id

    print("=" * 84)
    print("远端目录树（只读，不会写入任何内容）")
    print("=" * 84)
    print(f"  空间   {name}   space_id={space_id}")
    print(f"  根     {root}")
    print()

    printed = 0

    def walk(pid: str, indent: int) -> None:
        nonlocal printed
        if indent > depth or printed >= limit:
            return
        try:
            kids = cli.list_children(pid)
        except LC.LexiangError as e:
            print("  " * indent + f"[WARN] {e}")
            return
        for ch in kids:
            if printed >= limit:
                print("  " * indent + "…… 已达显示上限")
                return
            eid = entry_id_of(ch)
            nm = str(_first(ch, ("name", "title"), "-"))
            if _is_folder(ch):
                print("  " * indent + f"[目录] {nm}   {eid}")
                printed += 1
                walk(eid, indent + 1)
            else:
                print("  " * indent + f"[文件] {nm}")

    walk(root, 1)
    print()
    print("  要镜像到某个目录，把它作为镜像根授权：")
    print("    python kbsync.py --kb-root <知识库根> --space-id %s \\" % space_id)
    print("                      --root-entry-id <上面的 entry_id> --authorize")
    return 0


def cmd_revoke(kb_root: Path, space_id: str | None) -> int:
    auth = AuthStore(kb_root)
    if not auth.domains:
        print("  当前没有任何授权域记录。")
        return 0
    target = space_id or auth.domains[0].get("space_id")
    if auth.drop(str(target)):
        auth.save()
        print(f"  已撤销授权域 {target}。远端内容不受影响（本脚本从不删除远端条目）。")
        return 0
    print(f"  未找到授权域 {target}")
    return 2


# ---------------------------------------------------------------------------
# 主同步
# ---------------------------------------------------------------------------


def inspect_remote(items: list[SyncItem], dom: dict, root_id: str,
                   aliases: dict, accept_conflicts: bool) -> int:
    """只读预演：连远端解析目录链 + 同源比对，**零写入**。

    回答的是"如果现在执行，会传到哪、哪些会被判定为已存在"。
    与 --apply 共用 `resolve_folder_ro` / `find_same_doc`，所以预演结果与执行结果同源，
    不会出现"预演说复用、执行却重传"的偏差。

    返回 0 正常 / 2 连接失败。
    """
    pending = [i for i in items if i.status == "pending"]
    if not pending:
        return 0
    try:
        cli = LC.LexiangMCP()
        cli.handshake()
        cli.profile()
    except LC.LexiangError as e:
        print(f"[FAIL] 连接乐享失败：{e}")
        return 2

    index = RemoteIndex(cli)
    cache: dict = dict(dom.get("folders") or {})
    used: set[str] = set()
    for it in pending:
        parent, miss = resolve_folder_ro(index, cache, root_id, it.parts, aliases)
        if miss:
            it.folder_state = "create"
            it.notes = (it.notes + f" · 目录 [{miss}] 尚不存在，执行时将补建").strip(" ·")
            continue
        it.folder_state = "exists"
        ent, eid, tier, reason = find_same_doc(index, parent, it.stem, used)
        if not ent:
            continue
        it.remote_entry_id = eid
        it.match_tier = tier
        it.match_reason = reason
        used.add(eid)                      # 一对一：认领后不再给别人
        if tier == "weak" and not accept_conflicts:
            it.status = "conflict"
            it.notes = (it.notes + " · 疑似同源但确信度不足，待人工确认").strip(" ·")
    return 0


def cmd_sync(args) -> int:
    kb_root = Path(args.kb_root)
    if not kb_root.exists():
        print(f"[FAIL] 知识库根不存在：{kb_root}")
        return 2

    auth = AuthStore(kb_root)
    dom = auth.get(args.space_id)
    if dom is None:
        print("=" * 84)
        print("拒绝同步：没有可用的授权域")
        print("=" * 84)
        if args.space_id:
            print(f"  space_id {args.space_id} 不在 {S.AUTH_NAME} 里。")
        print("  必须先由你显式授权一个知识库空间 —— 这是官方写入红线的要求：")
        print("  禁止未确认就写入，禁止自行挑选目标。")
        print()
        print("  python kbsync.py --kb-root \"%s\" --space-id <space_id> --authorize" % kb_root)
        return 3
    root_id = str(dom.get("root_entry_id") or "")
    if not root_id:
        print("[FAIL] 授权域缺少 root_entry_id，请重新授权。")
        return 3

    skill_dir = Path(args.skill_dir) if args.skill_dir else _HERE.parent
    tax = load_taxonomy(skill_dir)
    vocab = tag_vocab_of(tax)
    aliases = space_aliases_of(tax)
    st = S.StateStore(kb_root)
    slog = S.SyncLog(kb_root)

    items = build_plan(kb_root, st, vocab, subpath=args.subpath,
                       force=args.force, limit=args.limit)

    # 只读预演：先连远端把目录/同源判定跑一遍，再渲染计划，让计划反映真实判定结果
    if args.inspect and not args.apply:
        rc = inspect_remote(items, dom, root_id, aliases, args.accept_conflicts)
        if rc:
            return rc

    print_plan(items, kb_root, dom, applied=args.apply,
               batch_size=args.batch_size, force=args.force,
               inspected=args.inspect)
    if args.inspect:
        print_conflicts(items)

    if args.json:
        Path(args.json).write_text(json.dumps({
            "kb_root": str(kb_root),
            "space_id": dom.get("space_id"),
            "root_entry_id": root_id,
            "items": [{
                "seq": i.seq, "rel": i.rel, "md5": i.md5, "size": i.size,
                "status": i.status, "doc_type": i.doc_type, "subject": i.subject,
                "tags": i.tags, "remote_entry_id": i.remote_entry_id,
                "remote_url": i.remote_url, "match_tier": i.match_tier,
                "match_reason": i.match_reason, "folder_state": i.folder_state,
                "notes": i.notes, "error": i.error,
            } for i in items],
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  计划已写入 {args.json}\n")

    pending = [i for i in items if i.status == "pending"]
    cf_n = sum(1 for i in items if i.status == "conflict")
    if not pending:
        print("  没有可直接执行的条目 —— 全部挂在『待确认』，先处理上面的清单。"
              if cf_n else "  没有需要同步的内容 —— 本地与远端台账一致。")
        print()
        return 0
    if not args.apply:
        print("  这是只读预演，不会写入任何内容。确认无误后加 --apply 执行。"
              if args.inspect else
              "  这是纯本地 dry-run。加 --inspect 可预演远端匹配，加 --apply 执行。")
        print()
        return 0

    # ---- 真正开始写 ----
    try:
        cli = LC.LexiangMCP()
        cli.handshake()
        prof = cli.profile()
    except LC.LexiangError as e:
        print(f"[FAIL] 连接乐享失败：{e}")
        return 2

    domain = str(prof.get("company_domain") or "")
    company_from = str(prof.get("company_code") or "")

    index = RemoteIndex(cli)
    cache: dict = dict(dom.get("folders") or {})
    batch_id = S.new_batch_id()
    operator = args.operator
    t_all = time.time()
    bsize = max(1, args.batch_size)
    used: set[str] = set()        # 已被认领的远端 entry_id，保证同源判定一对一

    print(f"  开始同步，批次 {batch_id}，共 {len(pending)} 份，每批 {bsize} 份")
    print("-" * 84)

    for bi in range(0, len(pending), bsize):
        batch = pending[bi:bi + bsize]
        rows: list[dict] = []
        for it in batch:
            t0 = time.time()
            folder = it.remote_dir
            try:
                parent_id, created, aliased = ensure_folder(
                    index, cache, root_id, it.parts, aliases)
                for c in created:
                    rows.append({
                        "ts": S.now_iso(), "batch_id": batch_id, "action": "create_folder",
                        "remote_folder": c, "remote_entry_id": cache.get(c, ""),
                        "operator": operator, "notes": "自动补建远端目录",
                    })
                for a in aliased:
                    rows.append({
                        "ts": S.now_iso(), "batch_id": batch_id, "action": "match_folder",
                        "remote_folder": a.split(" → ")[0], "remote_entry_id": cache.get(a.split(" → ")[0], ""),
                        "operator": operator, "notes": a,
                    })
                if aliased:
                    it.notes = (it.notes + " · " + "；".join(aliased)).strip(" ·")
                it.remote_folder = folder

                # 用「文档身份」而非「文件名」判断远端是否已有同源条目
                # —— 远端文件名与本地并不逐一对应（日期被改、标题被缩写、格式被转）
                ent, eid, tier, reason = find_same_doc(index, parent_id, it.stem, used)
                if ent and not args.force:
                    it.match_tier = tier
                    it.match_reason = reason
                    if tier == "weak" and not args.accept_conflicts:
                        # 疑似同源但不够确信 —— 挂起交人，且**不写台账**，
                        # 这样下次 --inspect 仍能看到它，不会静默漏掉一份文件
                        it.status = "conflict"
                        it.remote_entry_id = eid
                        it.notes = (it.notes + " · 疑似同源但确信度不足，待人工确认").strip(" ·")
                    else:
                        used.add(eid)                 # 一对一：认领后不再给别人
                        it.remote_entry_id = eid
                        it.remote_url = build_url(domain, company_from, eid)
                        it.status = "reused"
                        it.notes = (it.notes + " · 远端已有同源条目，复用不重传").strip(" ·")
                        if args.tags and it.tags:
                            try:
                                cli.set_tags(eid, it.tags)
                            except Exception as e:
                                it.notes += f" · 标签写入失败：{e}"
                        st.put(it.md5,
                               remote_entry_id=it.remote_entry_id,
                               remote_url=it.remote_url,
                               remote_space_id=dom.get("space_id"),
                               synced_at=S.now_iso(),
                               sync_batch=batch_id,
                               status="synced")
                else:
                    upload_one(cli, it, parent_id, args.tags, domain, company_from)
                    it.status = "synced"
                    used.add(it.remote_entry_id)
                    index.invalidate(parent_id)          # 目录内容变了
                    st.put(it.md5,
                           remote_entry_id=it.remote_entry_id,
                           remote_url=it.remote_url,
                           remote_space_id=dom.get("space_id"),
                           synced_at=S.now_iso(),
                           sync_batch=batch_id,
                           status="synced")
                it.elapsed = time.time() - t0
            except Exception as e:      # 单份失败不拖垮整批
                it.status = "failed"
                it.error = f"{type(e).__name__}: {e}"
                it.elapsed = time.time() - t0

            rows.append({
                "ts": S.now_iso(), "batch_id": batch_id, "md5": it.md5,
                "doc_type": it.doc_type, "subject": it.subject,
                "target_rel": it.rel, "remote_folder": it.remote_folder,
                "action": {"synced": "upload", "reused": "reuse",
                           "conflict": "conflict", "failed": "error",
                           "skipped": "skip"}.get(it.status, it.status),
                "remote_entry_id": it.remote_entry_id, "remote_url": it.remote_url,
                "size": it.size, "elapsed_s": f"{it.elapsed:.2f}",
                "tags": ",".join(it.tags), "operator": operator,
                "notes": " | ".join(filter(None, [it.notes, it.error]))[:800],
            })

        # 每批落一次盘：中途崩溃也能续传
        st.save()
        auth.set_folders(str(dom.get("space_id")), cache)
        auth.save()
        slog.append(rows)

        done = min(bi + bsize, len(pending))
        ok_n = sum(1 for i in batch if i.status in ("synced", "reused"))
        print(f"  批次进度 {done}/{len(pending)}  本批成功 {ok_n}/{len(batch)}  "
              f"累计 {fmt_size(sum(i.size for i in pending[:done] if i.status == 'synced'))}")

    print("-" * 84)
    print_conflicts(items)
    ok, skip, fail = print_result(items, batch_id, time.time() - t_all)
    return 0 if fail == 0 else 4


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        prog="kbsync",
        description="乐享知识库增量同步引擎（默认 dry-run，需显式授权域）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="首次使用顺序：\n"
               "  1) python lexiang_client.py                      # 自检连通性\n"
               "  2) python kbsync.py --discover                   # 只读查看可用空间\n"
               "  3) python kbsync.py --kb-root <根> --space-id <id> --authorize\n"
               "  4) python kbsync.py --kb-root <根>               # 看本地计划（零网络）\n"
               "  5) python kbsync.py --kb-root <根> --inspect     # 连远端预演（零写入）\n"
               "  6) python kbsync.py --kb-root <根> --apply       # 执行\n")
    ap.add_argument("--kb-root", default=os.environ.get("KB_ROOT", ""),
                    help="本地知识库根目录（也支持环境变量 KB_ROOT）")
    ap.add_argument("--skill-dir", default="",
                    help="skill 根目录（用于读取 assets/taxonomy.yaml 的标签词表）")
    ap.add_argument("--space-id", default="", help="乐享知识库 space_id")
    ap.add_argument("--root-entry-id", default="",
                    help="镜像根：授权到空间内某个子目录而非空间根（--authorize 时生效）")
    ap.add_argument("--skill-config", dest="config", default="",
                    help="凭据 JSON 路径（默认自动从本机连接器发现）")

    g = ap.add_mutually_exclusive_group()
    g.add_argument("--authorize", action="store_true", help="把 --space-id 登记为授权域")
    g.add_argument("--revoke", action="store_true", help="撤销授权域")
    g.add_argument("--discover", action="store_true", help="只读列出可见团队/知识库")
    g.add_argument("--tree", action="store_true", help="只读打印某空间的目录树")
    g.add_argument("--show-auth", action="store_true", help="查看授权域记录")

    ap.add_argument("--depth", type=int, default=2, help="--tree 的递归深度")

    ap.add_argument("--apply", action="store_true", help="真正执行（不加则为 dry-run）")
    ap.add_argument("--inspect", action="store_true",
                    help="只读预演：连远端解析目录链与同源条目，零写入（可与 --json 搭配出预演报告）")
    ap.add_argument("--accept-conflicts", dest="accept_conflicts", action="store_true",
                    help="把『疑似同源』(weak 档) 也认定为同一份并复用（默认挂起交人确认）")
    ap.add_argument("--force", action="store_true",
                    help="忽略台账与同源判定，强制重传（默认关闭，只创建不覆盖）")
    ap.add_argument("--tags", dest="tags", action="store_true", default=True,
                    help="写入标签（默认开启）")
    ap.add_argument("--no-tags", dest="tags", action="store_false", help="不写标签")
    ap.add_argument("--limit", type=int, default=0, help="本轮最多同步多少份（试水用）")
    ap.add_argument("--subpath", default="", help="只同步某个子树（相对知识库根）")
    ap.add_argument("--batch-size", type=int, default=DEFAULT_BATCH_SIZE,
                    help=f"每批份数，官方规范上限 20（默认 {DEFAULT_BATCH_SIZE}）")
    ap.add_argument("--operator", default="agent", help="执行者标识，记入审计日志")
    ap.add_argument("--notes", default="", help="授权备注")
    ap.add_argument("--json", default="", help="把计划写到指定 JSON 文件")
    args = ap.parse_args(argv)

    # 显式给到的凭据路径优先 —— 写进环境变量让 lexiang_client 的统一入口读到
    if args.config:
        os.environ["KB_LEXIANG_CONFIG"] = args.config

    if args.discover:
        return cmd_discover()
    if args.tree:
        if not args.space_id:
            print("[FAIL] --tree 需要 --space-id")
            return 2
        return cmd_tree(args.space_id, depth=args.depth, root_entry_id=args.root_entry_id)
    if not args.kb_root:
        ap.print_help()
        print("\n[FAIL] 缺少 --kb-root")
        return 2

    kb_root = Path(args.kb_root)
    if args.show_auth:
        return cmd_show_auth(kb_root, as_json=False)
    if args.authorize:
        return cmd_authorize(kb_root, args.space_id, args.operator, args.notes,
                             args.config or None, args.root_entry_id)
    if args.revoke:
        return cmd_revoke(kb_root, args.space_id or None)

    try:
        return cmd_sync(args)
    except LC.AuthError as e:
        print(f"\n[FAIL] {e}")
        return 2
    except KeyboardInterrupt:
        print("\n[中断] 已完成部分已逐批落盘，直接重跑即可续传。")
        return 4


if __name__ == "__main__":
    raise SystemExit(main())
