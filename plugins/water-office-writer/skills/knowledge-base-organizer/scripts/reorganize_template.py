"""临时一次性脚本：彻底重组知识库副本

用法：
    python reorganize.py              # 全量重建（先清理旧输出）
    python reorganize.py --no-clean   # 原地刷新（复用已存在文件，只重写索引与元数据）
"""
import csv, shutil, sys
from pathlib import Path
from dataclasses import dataclass, field
from typing import List

SRC = Path(r"E:/深圳水务集团/2.2-数字员工项目/知识库")
OUT = SRC / "_reorg_20260910"

# 主空间与子目录的业务排序（董事长优先，其后按"总结类 → 对外类 → 专题类"）
SPACE_ORDER = ["00-规范制度", "01-龚利民（董事长）", "02-方琳（总裁）"]
SUB_ORDER = [
    "集团年度工作总结", "集团半年工作总结",
    "重要会议致辞", "对外汇报（部委）", "专题汇报", "就职演讲",
]


def _rank(name, order):
    return order.index(name) if name in order else len(order)


@dataclass
class Item:
    src_rel: str
    target_dir: str
    type: str
    date: str
    title_clean: str
    leader: str
    audience: str
    security: str
    tags: List[str] = field(default_factory=list)
    notes: str = ""

    def new_filename(self):
        d = self.date.replace("-", "")
        tag = "无领导" if self.leader == "-" else self.leader
        title = self.title_clean.replace("/", "-").replace(":", "-").replace("：", "-").strip(" -_")
        idx = self.src_rel.rfind(".")
        ext = self.src_rel[idx:] if idx >= 0 else ""
        return f"{d}_{self.type}_{title}_{tag}{ext}"


ITEMS = [
    Item("./规范.docx", "00-规范制度", "规范", "2024-01-01", "公文格式规范",
         "-", "全员", "内部", ["规范", "格式", "公文"], "集团内部公文格式标准"),

    Item("./方琳-总裁/【方总2025年工作总结】【20260211】方琳（第三次审稿后）20260209在集团2025年终总结大会上的讲话(2).docx",
         "02-方琳（总裁）/集团年度工作总结", "年终总结", "2026-02-11",
         "集团2025年终总结大会讲话", "方琳", "全员", "内部", ["年终总结", "2025", "方琳", "集团"]),
    Item("./方琳-总裁/【方总2025年半年工作总结】【20250715】会后完善定稿0710在集团2025半年工作总结大会上的讲话（方琳总）.docx",
         "02-方琳（总裁）/集团半年工作总结", "半年总结", "2025-07-15",
         "集团2025半年总结大会讲话", "方琳", "全员", "内部", ["半年总结", "2025", "方琳", "集团"]),
    Item("./方琳-总裁/【20260723】【终稿】0723-2集团2026年上半年工作总结及下半年工作部署（方琳总）.pptx",
         "02-方琳（总裁）/集团半年工作总结", "半年总结", "2026-07-23",
         "集团2026上半年总结及下半年部署", "方琳", "全员", "内部",
         ["半年总结", "2026", "方琳", "集团", "PPT"], "大文件 79.5MB"),

    Item("./龚利民-董事长/【集团2023年的工作总结】（终稿已更新全年数据）深圳环境水务集团关于2023年度工作总结及2024年度工作计划的报告.docx",
         "01-龚利民（董事长）/集团年度工作总结", "年终总结", "2024-01-01",
         "集团2023年度工作总结及2024年度工作计划报告", "龚利民", "全员", "内部",
         ["年终总结", "2023", "龚利民", "集团"]),
    Item("./龚利民-董事长/【集团2024年的工作总结】【20250123】深圳环境水务集团2024年工作总结和2025年工作计划.doc",
         "01-龚利民（董事长）/集团年度工作总结", "年终总结", "2025-01-23",
         "集团2024年工作总结和2025年工作计划", "龚利民", "全员", "内部",
         ["年终总结", "2024", "龚利民", "集团"], "旧版.doc，建议转.docx"),
    Item("./龚利民-董事长/【集团2025年的工作总结】附件1 深圳环境水务集团2025年度工作总结及2026年度工作计划.doc",
         "01-龚利民（董事长）/集团年度工作总结", "年终总结", "2026-01-01",
         "集团2025年度工作总结及2026年度工作计划", "龚利民", "全员", "内部",
         ["年终总结", "2025", "龚利民", "集团"], "旧版.doc，建议转.docx"),
    Item("./龚利民-董事长/【龚董2024年工作总结】【20250123】大变局下的战略突破（龚总演讲稿）.docx",
         "01-龚利民（董事长）/集团年度工作总结", "年终总结", "2025-01-23",
         "大变局下的战略突破", "龚利民", "全员", "内部",
         ["年终总结", "2024", "龚利民", "个人总结", "战略"]),
    Item("./龚利民-董事长/【龚董2025年工作总结】【20260211】龚董20260210（23点）在集团2025年终工作总结大会上的讲话(1).docx",
         "01-龚利民（董事长）/集团年度工作总结", "年终总结", "2026-02-11",
         "集团2025年终工作总结大会讲话", "龚利民", "全员", "内部",
         ["年终总结", "2025", "龚利民", "集团"]),

    Item("./龚利民-董事长/【龚董2025年半年工作总结】【20250711】0711 1300（董事长）在集团2025半年工作总结会议上的讲话.docx",
         "01-龚利民（董事长）/集团半年工作总结", "半年总结", "2025-07-11",
         "集团2025半年总结会议讲话", "龚利民", "全员", "内部",
         ["半年总结", "2025", "龚利民", "集团"]),
    Item("./龚利民-董事长/【龚董2026年半年总结PPT】【最终版】以优良作风践行正确政绩观 推动集团高质量发展迈上新台阶（龚董）.pdf",
         "01-龚利民（董事长）/集团半年工作总结", "半年总结", "2026-08-01",
         "以优良作风践行正确政绩观 推动集团高质量发展迈上新台阶",
         "龚利民", "全员", "内部",
         ["半年总结", "2026", "龚利民", "集团", "政绩观"]),

    Item("./龚利民-董事长/【2026年中国水协年会】主旨报告.pptx",
         "01-龚利民（董事长）/重要会议致辞", "主旨报告", "2026-01-01",
         "中国水协年会主旨报告", "龚利民", "行业协会", "公开",
         ["致辞", "2026", "龚利民", "水协", "年会", "PPT"],
         "大文件 439.7MB，建议压缩或转PDF"),
    Item("./龚利民-董事长/【产业科技博览会上的致辞】【202511】（龚董）在2025环境水务产业科技博览会上的致辞.docx",
         "01-龚利民（董事长）/重要会议致辞", "致辞", "2025-11-01",
         "2025环境水务产业科技博览会致辞", "龚利民", "行业协会", "公开",
         ["致辞", "2025", "龚利民", "博览会", "科技"]),
    Item("./龚利民-董事长/【粤港澳大湾区论坛】【2024】在第六届粤港澳大湾区水务论坛暨第十五届深港珠澳供水届学术交流会上的致辞 (1)(2).docx",
         "01-龚利民（董事长）/重要会议致辞", "致辞", "2024-01-01",
         "第六届粤港澳大湾区水务论坛致辞", "龚利民", "行业协会", "公开",
         ["致辞", "2024", "龚利民", "粤港澳大湾区"]),
    Item("./龚利民-董事长/【深圳市供排水协会会员大会】【2024】在深圳市供排水行业协会第三届会员大会上的致辞（龚总）.doc",
         "01-龚利民（董事长）/重要会议致辞", "致辞", "2024-01-01",
         "深圳市供排水行业协会第三届会员大会致辞", "龚利民", "行业协会", "公开",
         ["致辞", "2024", "龚利民", "深圳供排水协会"], "旧版.doc，建议转.docx"),
    Item("./龚利民-董事长/【省人工智能对接大会】【202604】在广东省人工智能应用对接大会的讲话0426.docx",
         "01-龚利民（董事长）/重要会议致辞", "讲话", "2026-04-01",
         "广东省人工智能应用对接大会讲话", "龚利民", "行业协会", "公开",
         ["致辞", "2026", "龚利民", "AI", "人工智能"]),

    Item("./龚利民-董事长/【报住建部】【2026】 （董事长拜访住建部领导）千家万户水管家：深圳超大城市水务服务新路径.docx",
         "01-龚利民（董事长）/对外汇报（部委）", "对外汇报", "2026-01-01",
         "千家万户水管家-深圳超大城市水务服务新路径", "龚利民", "部委", "敏感",
         ["对外汇报", "2026", "龚利民", "住建部"]),
    Item("./龚利民-董事长/【报水利部】深圳环境水务集团关于近期主要工作情况的报告.docx",
         "01-龚利民（董事长）/对外汇报（部委）", "对外汇报", "2026-01-01",
         "近期主要工作情况的报告-水利部", "龚利民", "部委", "敏感",
         ["对外汇报", "2026", "龚利民", "水利部"]),
    Item("./龚利民-董事长/【报生态环境部】深圳环境水务集团关于近期主要工作情况的报告.docx",
         "01-龚利民（董事长）/对外汇报（部委）", "对外汇报", "2026-01-01",
         "近期主要工作情况的报告-生态环境部", "龚利民", "部委", "敏感",
         ["对外汇报", "2026", "龚利民", "生态环境部"]),

    Item("./龚利民-董事长/【鸿蒙】【20260818】发市长0818（定稿）深圳环境水务集团关于构建水务鸿蒙生态情况的报告.docx",
         "01-龚利民（董事长）/专题汇报", "专题汇报", "2026-08-18",
         "关于构建水务鸿蒙生态情况的报告", "龚利民", "市政府", "敏感",
         ["专题汇报", "2026", "龚利民", "鸿蒙", "智慧水务"]),

    Item("./龚利民-董事长/【水协科技委会长】【2024】就职演讲（龚利民）.docx",
         "01-龚利民（董事长）/就职演讲", "就职演讲", "2024-01-01",
         "水协科技委会长就职演讲", "龚利民", "行业协会", "公开",
         ["就职", "2024", "龚利民", "水协科技委"]),

    Item("./龚利民-董事长/【鸿蒙生态大会】【2026】20260828（董事长）在2026GIIC联盟鸿蒙生态大会水务鸿蒙生态论坛上的致辞(3).docx",
         "01-龚利民（董事长）/重要会议致辞", "致辞", "2026-08-28",
         "2026GIIC鸿蒙生态大会水务鸿蒙生态论坛致辞", "龚利民", "行业协会", "公开",
         ["致辞", "2026", "龚利民", "鸿蒙", "生态大会"]),
]


def copy_one(item):
    src = SRC / item.src_rel.lstrip("./").replace("/", "\\")
    if not src.exists():
        return None, f"源不存在: {src}"
    out_dir = OUT / item.target_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    new_name = item.new_filename()
    dst = out_dir / new_name
    reused = False
    if dst.exists() and dst.stat().st_size == src.stat().st_size:
        # 幂等：同名同大小，直接复用（避免被占用的大文件重复复制）
        reused = True
    else:
        if dst.exists():
            stem, suf = dst.stem, dst.suffix
            i = 2
            while True:
                cand = out_dir / f"{stem}_v{i}{suf}"
                if not cand.exists():
                    dst, new_name = cand, cand.name
                    break
                i += 1
        shutil.copy2(src, dst)
    return {
        "src_rel": item.src_rel,
        "src_size_mb": round(src.stat().st_size / 1024 / 1024, 3),
        "new_path": str(dst.relative_to(OUT)).replace("\\", "/"),
        "new_filename": new_name,
        "type": item.type, "date": item.date,
        "title_clean": item.title_clean, "leader": item.leader,
        "audience": item.audience, "security": item.security,
        "tags": ",".join(item.tags), "notes": item.notes,
    }, None, reused


def write_readmes(rows):
    spaces = {
        "00-规范制度": ("00 · 规范制度", "集团内部公文/报告写作规范与制度文件"),
        "01-龚利民（董事长）": ("01 · 龚利民（董事长）", "董事长龚利民的全部工作文件（内部与对外）"),
        "02-方琳（总裁）": ("02 · 方琳（总裁）", "总裁方琳的对外公开/对内工作总结与讲话"),
    }
    sub_intros = {
        "集团年度工作总结": "集团年度工作总结与工作计划（年终总结大会讲话）",
        "集团半年工作总结": "集团半年工作总结与下半年工作部署",
        "重要会议致辞": "董事长出席的重要行业/外部会议的致辞与主旨报告",
        "对外汇报（部委）": "对国家部委（住建部/水利部/生态环境部等）的专题汇报材料",
        "专题汇报": "面向特定对象（市/省/特定主题）的专题汇报",
        "就职演讲": "领导在行业协会/学会等机构担任职务的就职演讲",
    }
    by_dir = {}
    for row in rows:
        np = row["new_path"]
        parts = np.split("/", 1)
        space = parts[0]
        sub = "(根)"
        if len(parts) > 1 and "/" in parts[1]:
            sub = parts[1].split("/", 1)[0]
        by_dir.setdefault(f"{space}/{sub}", []).append(row)

    for space_key, (title, desc) in spaces.items():
        sub_keys = sorted([k for k in by_dir if k.startswith(space_key + "/")])
        root_key = f"{space_key}/(根)"
        total = sum(len(by_dir[k]) for k in sub_keys)
        actual_subs = sorted([k.split("/", 1)[1] for k in sub_keys if k != root_key],
                             key=lambda n: _rank(n, SUB_ORDER))
        lines = [f"# {title}", "", f"> {desc}", "",
                 f"**文件总数**：{total}  |  **子目录数**：{len(actual_subs)}", ""]
        if root_key in by_dir:
            lines += ["## 文件清单（直接存放在本目录）", "",
                      "| 日期 | 类型 | 标题 | 受众 | 密级 | 领导 | 文件 |",
                      "|------|------|------|------|------|------|------|"]
            for r in sorted(by_dir[root_key], key=lambda x: x["date"], reverse=True):
                lines.append(f"| {r['date']} | {r['type']} | {r['title_clean']} | {r['audience']} | {r['security']} | {r['leader']} | `{r['new_filename']}` |")
            lines += ["", "## 子目录索引", ""]
        else:
            lines += ["## 子目录索引", ""]
        for sub in actual_subs:
            intro = sub_intros.get(sub, "")
            lines.append(f"- **{sub}**（{len(by_dir[f'{space_key}/{sub}'])} 份）{('— ' + intro) if intro else ''}")
        lines += ["", "---", ""]
        (OUT / space_key).mkdir(parents=True, exist_ok=True)
        (OUT / space_key / "README.md").write_text("\n".join(lines), encoding="utf-8")

    for dir_key, dir_rows in by_dir.items():
        _, sub = dir_key.split("/", 1)
        if sub == "(根)":
            continue
        dir_path = OUT / dir_key
        dir_path.mkdir(parents=True, exist_ok=True)
        intro = sub_intros.get(sub, "本目录文件清单")
        lines = [f"# {sub}", "", f"> {intro}", "", f"**文件数**：{len(dir_rows)}", "",
                 "## 文件清单", "",
                 "| 日期 | 类型 | 标题 | 受众 | 密级 | 领导 | 文件 |",
                 "|------|------|------|------|------|------|------|"]
        for r in sorted(dir_rows, key=lambda x: (x["date"], x["title_clean"]), reverse=True):
            title = r["title_clean"]
            if r["notes"]:
                title += f" *（{r['notes']}）*"
            lines.append(f"| {r['date']} | {r['type']} | {title} | {r['audience']} | {r['security']} | {r['leader']} | `{r['new_filename']}` |")
        lines += ["", "---", f"*生成时间：2026-09-10  ·  共 {len(dir_rows)} 份*", ""]
        (dir_path / "README.md").write_text("\n".join(lines), encoding="utf-8")


def clean_old_output():
    """清理上一轮生成的输出（保留 _build），被进程占用的文件跳过"""
    if not OUT.exists():
        return
    for p in sorted(OUT.iterdir()):
        if p.name == "_build":
            continue
        try:
            if p.is_dir():
                shutil.rmtree(p, ignore_errors=True)
            else:
                p.unlink()
        except OSError:
            pass
    # 由内向外清理残留空目录（被占用文件的父目录会保留）
    for p in sorted(OUT.rglob("*"), key=lambda x: len(x.parts), reverse=True):
        if p.is_dir() and "_build" not in p.parts:
            try:
                p.rmdir()
            except OSError:
                pass


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    if "--no-clean" not in sys.argv:
        clean_old_output()
    print(f"=== 开始重组，共 {len(ITEMS)} 份 ===")
    rows = []
    n_reused = 0
    for i, item in enumerate(ITEMS, 1):
        row, err, reused = copy_one(item)
        if err:
            print(f"[{i:02d}] FAIL {err}")
        else:
            rows.append(row)
            flag = "REUSE" if reused else "OK   "
            if reused:
                n_reused += 1
            print(f"[{i:02d}] {flag} {row['new_filename']}  ({row['src_size_mb']} MB)")

    # 按"主空间 → 子目录 → 日期"的业务顺序排序元数据
    def sort_key(r):
        parts = r["new_path"].split("/")
        sub = parts[1] if len(parts) > 2 else ""
        return (_rank(parts[0], SPACE_ORDER), _rank(sub, SUB_ORDER), r["date"], r["title_clean"])

    rows.sort(key=sort_key)

    csv_path = OUT / "_metadata.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    print(f"\n[OK] CSV:  {csv_path}")

    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = Workbook()
        ws = wb.active
        ws.title = "知识库元数据"
        headers = list(rows[0].keys())
        ws.append(headers)
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="305496")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in rows:
            ws.append([row[h] for h in headers])
        for col_idx, h in enumerate(headers, 1):
            width = 60 if h == "new_path" else 45 if h == "title_clean" else 30 if h in ("src_rel", "new_filename", "tags", "notes") else 12
            ws.column_dimensions[chr(64 + col_idx)].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        xlsx_path = OUT / "_metadata.xlsx"
        wb.save(xlsx_path)
        print(f"[OK] Excel: {xlsx_path}")
    except ImportError:
        pass

    write_readmes(rows)
    print("[OK] README 已生成")

    # 清理 Office 临时锁文件（~$*），被占用则跳过
    n_clean = 0
    for p in OUT.rglob("~$*"):
        try:
            p.unlink()
            n_clean += 1
        except OSError:
            pass
    print(f"[OK] 清理 Office 锁文件 {n_clean} 个")

    # 残留物报告
    leftover_locks = [p for p in OUT.rglob("~$*") if "_build" not in p.parts]
    if leftover_locks:
        print("[WARN] 以下 Office 锁文件被进程占用未能清理（不影响使用，重启后可删）：")
        for p in leftover_locks:
            print(f"       {p.relative_to(OUT)}")

    print(f"\n=== 完成 ===  成功 {len(rows)}/{len(ITEMS)}  复用 {n_reused}")


if __name__ == "__main__":
    main()
