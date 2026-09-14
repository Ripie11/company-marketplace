#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""lexiang_client.py —— 乐享 MCP 最小客户端

为什么自己实现一个：
  乐享 MCP 是标准的 streamable-HTTP + JSON-RPC 2.0 端点，用 30 行 requests
  就能打通。自己实现的好处是 `kbsync.py` 可以**独立运行**（不必依赖宿主 App
  的 MCP 会话），从而支持批量、幂等、可重跑的同步作业。

协议要点（实测于 https://mcp.lexiang-app.com/mcp）：
  1. `initialize` 握手 → 服务端**不返回** Mcp-Session-Id（无状态模式），
     因此后续每个请求都独立携带 Authorization，无需维持会话
  2. 响应可能是 `application/json`（实测）或 `text/event-stream`，两种都要能读
  3. `tools/call` 的结果包在 `result.content[0].text` 里，且通常是一段 JSON 字符串，
     需要再解一层

凭据解析优先级（高 → 低）：
  ① 显式传参 → ② 环境变量 → ③ 本机 WorkBuddy 连接器配置（自动发现）

③ 的自动发现逻辑：读 `~/.workbuddy/connectors/*/mcp.json`，找 key 里含
`lexiang` 的条目拿 `url`；再读同目录 `connector-states.json` 的
`headerOverrides` 拿最新的 Authorization（连接器续期后 token 会更新到这里）。
这样**用户无需把 token 抄进任何项目文件** —— 凭据始终留在它原本该在的地方。
"""
from __future__ import annotations

import io
import json
import os
import sys
import time
from pathlib import Path
from typing import Any

import requests

if sys.stdout.encoding is None or sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    except Exception:
        pass

DEFAULT_URL = "https://mcp.lexiang-app.com/mcp"
CONNECTOR_GLOB = "connectors/*/mcp.json"


class LexiangError(RuntimeError):
    def __init__(self, message: str, code: int | None = None, raw: Any = None):
        super().__init__(message)
        self.code = code
        self.raw = raw


class AuthError(LexiangError):
    """401 / 鉴权失败 —— 调用方应当引导用户重新授权，而不是重试。"""


# ---------------------------------------------------------------------------
# 凭据发现
# ---------------------------------------------------------------------------


def discover_credentials() -> tuple[str, str, str]:
    """返回 (url, token, 来源说明)。找不到 token 则抛 AuthError。"""
    # ① 项目级显式配置文件（可选，便于交付给同事时统一注入）
    cfg = Path(os.environ.get("KB_LEXIANG_CONFIG", "")) if os.environ.get("KB_LEXIANG_CONFIG") else None
    if cfg and cfg.exists():
        d = json.loads(cfg.read_text(encoding="utf-8"))
        if d.get("url") and d.get("token"):
            return d["url"], d["token"], f"配置文件 {cfg}"

    # ② 环境变量
    url = os.environ.get("LEXIANG_MCP_URL") or DEFAULT_URL
    tok = os.environ.get("LEXIANG_TOKEN")
    if tok:
        return url, tok, "环境变量 LEXIANG_TOKEN"

    # ③ 本机连接器自动发现
    home = Path(os.environ.get("USERPROFILE") or os.path.expanduser("~"))
    for mcp_json in sorted((home / ".workbuddy").glob(CONNECTOR_GLOB), reverse=True):
        try:
            data = json.loads(mcp_json.read_text(encoding="utf-8"))
        except Exception:
            continue
        servers = data.get("mcpServers") or {}
        for key, val in servers.items():
            if "lexiang" not in key.lower():
                continue
            url = val.get("url") or DEFAULT_URL
            token = (val.get("headers") or {}).get("Authorization") or ""
            # 连接器状态文件里的 headerOverrides 是续期后的最新值，优先级更高
            states = mcp_json.parent / "connector-states.json"
            if states.exists():
                try:
                    sd = json.loads(states.read_text(encoding="utf-8"))
                    ov = (sd.get("headerOverrides") or {}).get(
                        key.split(":", 1)[-1]) or (sd.get("headerOverrides") or {}).get(key) or {}
                    token = ov.get("Authorization") or token
                    disabled = key.split(":", 1)[-1] not in (sd.get("enabled") or [None])
                    if disabled and (sd.get("enabled") or []):
                        continue
                except Exception:
                    pass
            if token:
                return url, token, f"本机连接器 {mcp_json.parent.name}"
    raise AuthError(
        "找不到乐享凭据。请任选一种方式：\n"
        "  · 在 WorkBuddy「连接器」页面授权乐享（推荐，token 自动写入）\n"
        "  · 设置环境变量 LEXIANG_TOKEN（可另设 LEXIANG_MCP_URL）\n"
        "  · 用 --config <json> 指定 {'url':..., 'token':...}"
    )


# ---------------------------------------------------------------------------
# 客户端
# ---------------------------------------------------------------------------


class LexiangMCP:
    """极简 MCP over streamable-HTTP 客户端（只实现本 skill 用到的部分）。"""

    def __init__(self, url: str | None = None, token: str | None = None,
                 timeout: int = 120, retries: int = 2, verbose: bool = False):
        if url and token:
            self.url, self.token, self.source = url, token, "显式传参"
        else:
            d_url, d_tok, d_src = discover_credentials()
            self.url, self.token, self.source = url or d_url, token or d_tok, d_src
        self.timeout = timeout
        self.retries = retries
        self.verbose = verbose
        self._rid = 0
        self._tools: dict[str, dict] | None = None
        self._handshaken = False

    # ---- 底层 ----

    def _headers(self) -> dict:
        return {
            "Authorization": self.token,
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
        }

    @staticmethod
    def _parse(resp) -> Any:
        ct = (resp.headers.get("Content-Type") or "").lower()
        if "event-stream" in ct:
            for line in resp.text.splitlines():
                if line.startswith("data:"):
                    payload = line[5:].strip()
                    if payload and payload != "[DONE]":
                        try:
                            return json.loads(payload)
                        except Exception:
                            continue
            return None
        try:
            return resp.json()
        except Exception:
            return {"_raw": resp.text[:800]}

    def _rpc(self, method: str, params: dict | None = None) -> Any:
        self._rid += 1
        body: dict[str, Any] = {"jsonrpc": "2.0", "id": self._rid, "method": method}
        if params is not None:
            body["params"] = params
        payload = json.dumps(body, ensure_ascii=False).encode("utf-8")

        last_err: Exception | None = None
        for attempt in range(self.retries + 1):
            try:
                r = requests.post(self.url, headers=self._headers(),
                                  data=payload, timeout=self.timeout)
            except requests.RequestException as e:
                last_err = e
                if attempt < self.retries:
                    time.sleep(1.5 * (attempt + 1))
                    continue
                raise LexiangError(f"网络请求失败（{type(e).__name__}: {e}）") from e

            if r.status_code == 401:
                raise AuthError(
                    "乐享返回 401 —— 令牌已过期。请到 WorkBuddy「连接器」页面"
                    "对乐享点『重新授权』，token 会自动写回，无需改任何配置。",
                    code=401)
            if r.status_code == 429:
                last_err = LexiangError("被限流（429）", code=429)
                time.sleep(3 * (attempt + 1))
                continue
            if r.status_code >= 500:
                last_err = LexiangError(f"服务端错误 {r.status_code}", code=r.status_code)
                if attempt < self.retries:
                    time.sleep(2 * (attempt + 1))
                    continue
                raise last_err
            if r.status_code >= 400:
                raise LexiangError(f"HTTP {r.status_code}：{r.text[:300]}", code=r.status_code)

            data = self._parse(r)
            if isinstance(data, dict) and data.get("error"):
                err = data["error"]
                raise LexiangError(f"{err.get('message') or err}", raw=err)
            return (data or {}).get("result")

        raise last_err or LexiangError("未知失败")

    # ---- 会话 ----

    def handshake(self) -> dict:
        if self._handshaken:
            return {}
        res = self._rpc("initialize", {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "knowledge-base-organizer", "version": "1.0"},
        })
        self._handshaken = True
        return res or {}

    # ---- 工具 ----

    def tool_schemas(self) -> dict[str, dict]:
        if self._tools is None:
            res = self._rpc("tools/list", {})
            self._tools = {t["name"]: t for t in (res or {}).get("tools", [])}
        return self._tools

    def call(self, name: str, arguments: dict | None = None) -> Any:
        """调用工具并拆掉 MCP 的两层包装，直接返回业务对象。"""
        res = self._rpc("tools/call", {"name": name, "arguments": arguments or {}})
        if res is None:
            return None
        if isinstance(res, dict) and res.get("isError"):
            texts = [c.get("text", "") for c in res.get("content", []) if isinstance(c, dict)]
            raise LexiangError(f"工具 {name} 返回错误：{' / '.join(texts)[:400]}")
        if isinstance(res, dict) and res.get("structuredContent") is not None:
            return res["structuredContent"]
        if isinstance(res, dict) and isinstance(res.get("content"), list):
            texts = [c.get("text", "") for c in res["content"]
                     if isinstance(c, dict) and c.get("text")]
            joined = "\n".join(texts)
            try:
                return json.loads(joined)
            except Exception:
                return {"text": joined}
        return res

    # ---- 业务便捷方法 ----

    def whoami(self) -> dict:
        """当前账号信息。真实结构（实测）：
        {
          "staff":   {id, display_name, organization, ...},
          "company": {id, code, name, company_domain, sub_server_type},
          "personal_space": {id, name, root_entry_id, space_type},
          "corp_staff_id": "LX001"
        }
        """
        return self.call("whoami") or {}

    def profile(self) -> dict:
        """把 whoami 拍平成好用的字段，兼容新旧两种结构。"""
        me = self.whoami()
        staff = me.get("staff") if isinstance(me.get("staff"), dict) else {}
        comp = me.get("company") if isinstance(me.get("company"), dict) else {}
        psp = me.get("personal_space") if isinstance(me.get("personal_space"), dict) else {}
        return {
            "display_name": staff.get("display_name") or staff.get("name")
                            or me.get("name") or me.get("english_name") or "-",
            "staff_id": staff.get("id") or me.get("staff_id") or "",
            "corp_staff_id": me.get("corp_staff_id") or "",
            "organization": staff.get("organization") or "",
            "company_name": comp.get("name") or "",
            "company_domain": comp.get("company_domain") or "",
            "company_code": comp.get("code") or "",
            "company_id": comp.get("id") or "",
            "personal_space_id": psp.get("id") or me.get("personal_space_id") or "",
            "personal_space_name": psp.get("name") or "",
            "personal_root_entry_id": psp.get("root_entry_id") or "",
        }

    def describe_space(self, space_id: str) -> dict:
        return self.call("space_describe_space", {"space_id": space_id}) or {}

    def list_teams(self) -> list[dict]:
        r = self.call("team_list_teams") or {}
        return r.get("teams") or r.get("items") or []

    def list_spaces(self, team_id: str) -> list[dict]:
        r = self.call("space_list_spaces", {"team_id": team_id}) or {}
        return r.get("spaces") or r.get("items") or []

    def list_children(self, parent_id: str, limit: int = 200) -> list[dict]:
        """翻页拉全量直接子节点。"""
        out: list[dict] = []
        token = None
        while True:
            args: dict[str, Any] = {"parent_id": parent_id, "limit": limit}
            if token:
                args["page_token"] = token
            r = self.call("entry_list_children", args) or {}
            batch = r.get("entries") or r.get("items") or []
            out.extend(batch)
            token = r.get("next_page_token") or r.get("page_token")
            if not token or not batch:
                break
        return out

    def create_entry(self, parent_id: str, name: str, entry_type: str = "folder",
                     after: str | None = None) -> dict:
        args = {"parent_entry_id": parent_id, "name": name, "entry_type": entry_type}
        if after:
            args["after"] = after
        r = self.call("entry_create_entry", args) or {}
        return r.get("entry") or r

    def describe_entry(self, entry_id: str) -> dict:
        r = self.call("entry_describe_entry", {"entry_id": entry_id}) or {}
        return r.get("entry") or r

    def apply_upload(self, parent_entry_id: str, name: str, size: int,
                     mime_type: str, file_id: str | None = None) -> dict:
        args = {
            "parent_entry_id": parent_entry_id,
            "name": name,
            "mime_type": mime_type,
            "size": size,
            "upload_type": "PRE_SIGNED_URL",
        }
        if file_id:
            args["file_id"] = file_id
        r = self.call("file_apply_upload", args) or {}
        return r.get("session") or r

    def commit_upload(self, session_id: str) -> dict:
        r = self.call("file_commit_upload", {"session_id": session_id}) or {}
        return r.get("entry") or r

    def set_tags(self, entry_id: str, add: list[str], delete: list[str] | None = None) -> Any:
        args: dict[str, Any] = {"entry_id": entry_id, "add_tags": add}
        if delete:
            args["del_tags"] = delete
        return self.call("knowledge_tag_set_entry_tags", args)

    @staticmethod
    def put_blob(upload_url: str, file_path: Path, mime_type: str,
                 timeout: int = 1800) -> None:
        """第 2 步：HTTP PUT 原始字节。

        ⚠️ 必须用 `data=<file object>`（流式），不能读进内存 ——
        知识库里有 80MB+ 的 pptx，读进内存再传容易直接把进程撑爆。
        """
        with open(file_path, "rb") as f:
            r = requests.put(upload_url, data=f,
                             headers={"Content-Type": mime_type}, timeout=timeout)
        if r.status_code not in (200, 201, 204):
            raise LexiangError(
                f"预签名上传失败 HTTP {r.status_code}：{r.text[:200]}")


# ---------------------------------------------------------------------------
# MIME 映射
# ---------------------------------------------------------------------------

MIME = {
    ".pdf": "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".doc": "application/msword",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls": "application/vnd.ms-excel",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".ppt": "application/vnd.ms-powerpoint",
    ".md": "text/markdown",
    ".txt": "text/plain",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".zip": "application/zip",
    ".json": "application/json",
}


def mime_of(path: Path) -> str:
    return MIME.get(path.suffix.lower(), "application/octet-stream")


# ---------------------------------------------------------------------------
# 自检
# ---------------------------------------------------------------------------


def doctor(verbose: bool = True) -> int:
    try:
        cli = LexiangMCP(verbose=verbose)
    except AuthError as e:
        print(f"[FAIL] {e}")
        return 2
    print(f"[OK] 凭据来源：{cli.source}")
    print(f"     endpoint：{cli.url}")
    try:
        cli.handshake()
        me = cli.whoami()
    except AuthError as e:
        print(f"[FAIL] {e}")
        return 2
    except LexiangError as e:
        print(f"[FAIL] {e}")
        return 2

    company = me.get("company") or {}
    print("\n" + "=" * 66)
    print("账号信息")
    print("=" * 66)
    for k, v in cli.profile().items():
        if v:
            print(f"  {k:<22} {v}")

    tools = cli.tool_schemas()
    print(f"\n  可用工具       {len(tools)} 个")
    need = ["entry_create_entry", "entry_list_children", "file_apply_upload",
            "file_commit_upload", "space_describe_space"]
    miss = [n for n in need if n not in tools]
    print(f"  同步所需工具   {'齐全' if not miss else '缺少 ' + ', '.join(miss)}")
    if verbose:
        print("\n" + "=" * 66)
        print("如何拿到授权域（space_id）")
        print("=" * 66)
        print("  ① 在乐享网页打开目标知识库，URL 形如")
        print("     https://<域名>/spaces/<space_id>")
        print("  ② 不确定是哪个？只读探查可见范围：")
        print("     python kbsync.py --discover")
        print("     只看某个空间已有的目录树：python kbsync.py --tree <space_id>")
        print("  ③ 授权（可指定镜像到某个子目录，而不是空间根）：")
        print("     python kbsync.py --kb-root <知识库根> --space-id <space_id> --authorize")
        print("     python kbsync.py --kb-root <知识库根> --space-id <space_id> \\")
        print("                       --root-entry-id <子目录entry_id> --authorize")
    return 0


if __name__ == "__main__":
    raise SystemExit(doctor())
