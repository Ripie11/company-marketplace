#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
regress.py — 报告生成链路全量回归运行器（技能自检）

流程：
    0. 若 selftest/cases 不存在，自动先跑 cases_gen.py 生成 10 份测试源稿
    1. 读 cases/_manifest.json，把每份 .md 按用例指定领导落版为 docx
    2. 调 verify_docx.py --batch 跑 13 组版式断言
    3. 汇总 PASS/FAIL，输出 JSON 明细；加 --audit 追加产物审计

用法（在 selftest/ 目录下执行）:
    python regress.py --out-dir out_v1 --json v1_verify.json [--pdf] [--audit]

退出码：0 = 全部通过；1 = 有失败断言。可直接用于改动前后对比或 CI。
"""
import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent      # .../zhibi-suiwen/selftest
SKILL = HERE.parent                          # .../zhibi-suiwen
PY = Path(sys.executable)


def _run(cmd, cwd):
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                          env={**os.environ, "PYTHONIOENCODING": "utf-8"},
                          cwd=str(cwd))


def ensure_cases() -> None:
    """用例不存在时自动生成"""
    if (HERE / "cases" / "_manifest.json").is_file():
        return
    print("[info] 未发现用例，先生成 …")
    r = _run([str(PY), str(HERE / "cases_gen.py")], HERE)
    if r.returncode != 0:
        raise SystemExit("[error] 用例生成失败：\n" + (r.stderr or r.stdout))
    print((r.stdout or "").strip()[:400])


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-dir", default="out_v1", help="产物输出子目录")
    ap.add_argument("--json", default="verify.json", help="校验明细 JSON")
    ap.add_argument("--pdf", action="store_true", help="对首份用例做 PDF 冒烟")
    ap.add_argument("--audit", action="store_true", help="追加产物审计")
    args = ap.parse_args()

    ensure_cases()
    out_dir = HERE / args.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    manifest = json.loads((HERE / "cases" / "_manifest.json").read_text(encoding="utf-8"))
    built, failed_build = [], []
    first_id = manifest[0]["case_id"]
    for m in manifest:
        cid = m["case_id"]
        md = Path(m["md_path"])
        if not md.is_file():
            cand = sorted((HERE / "cases").glob(f"{cid}_*.md"))
            if not cand:
                failed_build.append((cid, "找不到源稿"))
                continue
            md = cand[0]
        leader = m.get("leader", "") or ""
        out = out_dir / f"{cid}_{md.stem.split('_', 1)[1]}.docx"
        cmd = [str(PY), str(SKILL / "scripts" / "md_to_gongwen_docx.py"),
               "--input", str(md), "--output", str(out)]
        if leader:
            cmd += ["--leader", leader]
        if args.pdf and cid == first_id:
            cmd += ["--pdf"]
        r = _run(cmd, SKILL / "scripts")
        if r.returncode != 0:
            failed_build.append((cid, (r.stderr or r.stdout or "").strip()[-300:]))
        else:
            built.append(cid)
            for ln in (r.stdout or "").splitlines():
                if ln.strip().startswith(("[warn]", "PDF")):
                    print(f"  {cid} {ln.strip()}")

    print(f"落版完成 {len(built)}/{len(manifest)}"
          + (f"，失败 {len(failed_build)}" if failed_build else ""))
    for cid, err in failed_build:
        print(f"  x {cid}: {err}")

    rc = _run([str(PY), str(HERE / "verify_docx.py"), "--batch", str(out_dir),
               "--manifest", str(HERE / "cases" / "_manifest.json"),
               "--out", str(HERE / args.json)], HERE)
    print(rc.stdout.rstrip())
    if rc.stderr.strip():
        print("[stderr]", rc.stderr.strip()[-400:])

    if args.audit:
        ar = _run([str(PY), str(HERE / "audit_outputs.py"), str(out_dir)], HERE)
        print("\n=== 产物审计 ===")
        print(ar.stdout.rstrip())

    return rc.returncode


if __name__ == "__main__":
    raise SystemExit(main())
