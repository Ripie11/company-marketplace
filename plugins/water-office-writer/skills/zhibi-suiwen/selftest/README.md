# selftest/ — 报告生成链路自检工具链

本目录是**本技能**自带的可重复回归测试套件，
用于在改动 `scripts/` 或 `assets/*.yaml` 后**防止版式静默回退**。

## 为什么需要它

历史上真实踩过的坑（都是"肉眼看不出来"的类型）：

| 坑 | 现象 | 后果 |
|----|------|------|
| A4 漏设 | `page_width/height` 未赋值 | 静默回退 Letter 21.6×27.9cm |
| 页脚查错部件 | 页码域在独立 page part，却去查 `doc.element.xml` | 误判"无页码"，漏检 |
| 无缩进段落 | `paragraph_format.first_line_indent` 为 `None` | 校验器 TypeError |
| 枚举字符串 | `str(WD_ALIGN_PARAGRAPH.CENTER)` = `"CENTER (1)"` | `endswith("CENTER")` 误判 |
| EMU 浮点误差 | 37.01mm vs 期望 37.0mm | 页边距误判 |

## 快速开始

```bash
cd selftest
python regress.py --out-dir out_v1 --json v1_verify.json --pdf --audit
```

首次运行会自动调 `cases_gen.py` 生成 10 份测试源稿。
**退出码 0 = 全部通过，1 = 有失败断言**，可直接接 CI。

## 测试矩阵（10 用例 × 2 套领导版式 × 3 种输出格式）

| 用例 | 报告类型 | 领导 | 埋设的边界条件 |
|------|---------|------|--------------|
| C01 | 年度工作总结 | 龚利民 | 标准四段式、多级标题、落款 |
| C02 | 半年工作总结 | 方琳 | 方琳四段式、`（一）`条款 |
| C03 | 重要会议致辞 | 龚利民 | 短篇、无一级标题、称呼语顶格 |
| C04 | 对外汇报（部委） | 方琳 | **Markdown 管道表格（5×4）** |
| C05 | 专题汇报 | 方琳 | 行内 `**加粗**`、`1.` 有序列表、条款引语 |
| C06 | 会议纪要 | 方琳 | 三段式、`>` 引用块 |
| C07 | 调研报告 | 龚利民 | **宽表格**、`（1）`三级层次、长陈述句以冒号结尾 |
| C08 | 述职报告 | 方琳 | 述职体例 |
| C09 | 讲话稿（短篇） | 龚利民 | 极短文本 |
| C10 | 边界用例 | **未指定** | 无 `--leader`、`####` 四级标题、`- ` 列表、**超长称呼语** |

## 13 组版式断言

1. 纸张 A4（21.0 × 29.7cm）
2. 页边距上下左右（mm，容差 0.5mm 抗 EMU 浮点误差）
3. 主标题字体 = 方正小标宋简体
4. 主标题字号 = 22pt
5. 主标题居中
6. 一级标题数量（按 md 源稿推导）
7. 一级标题字体 = 黑体
8. 一级标题字号 = 正文级
9. 一级标题缩进 = 2 字符
10. 正文字体 = 仿宋_GB2312 / 字号 / 行距固定 28pt / 首行缩进与字号联动
11. 称呼语顶格（缩进 0）+ 条款引语楷体 + 落款右对齐
12. **页脚 PAGE 域**（遍历 `doc.part.package.iter_parts()`）
13. **Markdown 标记零残留**：`**` / `|` / `#` / `- ` / `*` 五类

## 文件说明

| 文件 | 作用 |
|------|------|
| `cases_gen.py` | 生成 10 份测试源稿到 `cases/`（含 `_manifest.json` 用例→领导映射） |
| `verify_docx.py` | 版式校验器，支持 `--docx` 单份 / `--batch` 批量 |
| `regress.py` | 全量回归运行器（落版 → 校验 → 汇总） |
| `audit_outputs.py` | 产物审计：纸张/表格格数/PAGE/标记残留一览表 |

## 三副本一致性（`../scripts/sync_copies.py`）

本技能同时存在于**三处**（插件包 P / 项目级 skill Q / 专家分发包 R），
插件外壳（`plugin.json`、`README.md`、`USER_GUIDE.md`、`agents/`、`avatars/`）
又存在于 P / R **两处**。改动任意一处后，务必跑：

```bash
# 巡检（有漂移退出码 1，可直接接 CI）
python ../scripts/sync_copies.py --check

# 同步：skill 层三副本 + 插件层两副本
python ../scripts/sync_copies.py --apply
```

有几个必须知道的约定：

| 约定 | 说明 |
|------|------|
| **管辖范围** | 按 `SKILLS` 注册表逐技能巡检，当前覆盖 `zhibi-suiwen` 与 `knowledge-base-organizer`。**新增技能必须登记**，否则它的漂移永远不会被发现 |
| **规范源** | 默认 **P**（插件包）。Q 的技能名是 `leadership-report-generator`，同步时自动做身份替换 |
| **版本分叉台账** | 少数文件以 Q / R 为准（P 侧为旧版或缺失），全部登记在脚本的 `TAKE_FROM` 常量里。**在 Q / R 侧修了 bug 却忘了登记，会被 P 侧旧版反向覆盖** |
| **改称留兼容轨** | 统一称谓（如「方琳（执行总裁）」→「方琳（总裁）」）后，`knowledge-base-organizer` 的 `aliases` / 旧→新迁移映射里**保留旧名**，那是给已归档历史素材的向后兼容，不要"顺手清理" |
| **插件层台账** | 同理，见 `PLUGIN_TAKE_FROM`；未裁决的差异放 `PLUGIN_UNRESOLVED`，只报告不自动同步 |
| **字节级 I/O** | 脚本只用 `read_bytes` / `write_bytes`。用 `write_text` 会把 CRLF 写成 `\r\r\n`，让 Markdown 表格静默失效（见下） |
| **行尾体检** | `--check` / `--apply` 都会扫 `\r\r\n`；发现即报，`--apply` 会自愈 |

> ⚠️ **`\r\r\n` 陷阱**：`str.splitlines()` 会把 `\r\r\n` 当成**两个**换行，
> 于是 Markdown 表格的行之间多出空行，`md_to_gongwen_docx.py` 的 `parse_md_table`
> 判定「表头下一行不是分隔行」而整块退化为普通段落 —— **落版静默失效，肉眼极难发现**。
> 2026-09-14 实际踩过一次（15 个文件）。回归判据是 `C04` / `C07` 用例的表格断言。

## 注意

- **`cases/` 属版本化契约**（纳入三副本同步）：`regress.py` 依赖它，各副本应能独立自测。
  由 `cases_gen.py` 生成，改动生成逻辑后需重新生成并跑同步。
- **`out_*/`、`*verify*.json` 属运行产物**（**不**纳入同步，勿提交）：
  由 `regress.py` 现场生成，删掉不影响任何契约。
- 修改 `md_to_gongwen_docx.py` 的 OOXML 操作时，务必用 `_insert_ordered()`：
  `tblPr` / `tcPr` 子元素必须按 schema 顺序插入，否则 Word 报"文档已损坏"。
- PDF 冒烟（`--pdf`）需 Windows + Microsoft Word；无 Word 时会打印 `[warn]` 但不影响 docx 断言。
