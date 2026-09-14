---
name: zhibi-suiwen
description: 综合办公室专用能力，用于学习领导（董事长、总裁等）的报告风格与格式规范，并按其风格生成新的工作报告、讲话稿、会议纪要、致辞、汇报材料等。当用户说"按龚董的风格写一份X报告"、"学习方琳总的讲话风格"、"按公司领导文风起草"、"模仿领导格式生成报告"、"基于知识库里的领导讲话生成新文稿"时立即触发。同时支持两类子任务：(1) **风格建模**——从知识库中读取多份历史样本（.docx / .pptx / .pdf / .md），解析其中语言风格（语气、句式、用词、结构、称呼、收尾、**金句**、词汇分层、禁用清单）与文档格式（字体、字号、行距、页边距、段落、标题层级、落款），生成结构化的"风格档案"与"语言手册（voice_<领导>.md）"；(2) **风格应用**——基于用户给出的新需求（主题、要点、大纲、字数），加载已有风格档案与语言手册，生成风格高度一致、格式完全合规的新报告（默认输出 .docx，可选 .md / .pdf），并遵守金句密度红线（总结 3–5 处 / 致辞 1–2 处 / 呈报 0–1 处）。本skill不联网、不调用外部API，纯本地处理，适合党政机关、国有企业、事业单位对公文文风与版式有严格规范的场景。
---

# zhibi-suiwen — 执笔随文

## 一、Skill 目标与适用边界

### 适用场景
- 综合办公室 / 党委办公室 / 总经理办公室日常撰写的：工作报告、年度总结、专题汇报、致辞、讲话稿、调研报告、会议纪要、述职报告。
- 用户明确给出"参考 XX 领导"、"按 XX 风格"、"模仿公司公文格式"、"沿用往期文风"等诉求。

### 不适用场景
- 文学创作、营销文案、新闻稿（非公文场景）。
- 学术论文（需要另一套格式体系，参考 GB/T 7713）。
- 法律合同、医疗文书等强规范行业文本。

## 一点五、路径约定（跨项目可移植）

本 skill **不绑定任何具体项目路径**。运行时按以下优先级解析路径：

| 用途 | 解析规则（高 → 低） |
|------|--------------------|
| **样本目录** | ① 用户显式指定的路径 → ② 项目记忆 `MEMORY.md` 中的「样本目录」约定 → ③ 主动向用户确认 |
| **格式档案** | ① 用户显式指定 → ② skill 内置 `assets/profile_<leader>_v*.yaml` → ③ 项目输出目录下的 `assets/` |
| **语言手册** | ① 用户显式指定 → ② skill 内置 `assets/voice_<leader>.md` → ③ 档案 `linguistic.voice` 块 → ④ `references/linguistic_patterns.md`（通用兜底） |
| **报告输出** | ① 用户显式指定 → ② 项目记忆 `MEMORY.md` 中的「输出目录」约定 → ③ 当前工作区根目录 |

**随包内置档案**（可直接使用，无需重新建模）：

| 类型 | 文件 | 说明 |
|------|------|------|
| 格式+语言（机读） | `assets/profile_龚利民_总结_v1.yaml` | 董事长龚利民 · 总结类 |
| 格式+语言（人读） | `assets/profile_龚利民_总结_v1.md` | 同上，人读版 |
| 格式+语言（机读） | `assets/profile_方琳_总结_v1.yaml` | 总裁方琳 · 总结类 |
| 格式+语言（人读） | `assets/profile_方琳_总结_v1.md` | 同上，人读版 |
| 文种专属 | `assets/profile_龚利民_致辞_v1.yaml/.md` | `--doc-type speech` 自动加载 |
| 文种专属 | `assets/profile_龚利民_呈报_v1.yaml/.md` | `--doc-type baogao` 自动加载 |
| 🎙️ **语言手册** | `assets/voice_龚利民.md` | **阶段 B 首选参照**：语气画像 / 33 条带出处金句 / 18 式句式 / 6 层词汇 / 禁用清单 |
| 🎙️ **语言手册** | `assets/voice_方琳.md` | 同上（31 条金句 / 14 式 / 6 层词汇 / 15 项比喻）⚠️ 语料仅 3 份，缺致辞与呈报语体 |

> ⚠️ `voice_<领导>.md` 是**语言维度的唯一权威**；`profile_*.yaml/.md` 的 §三/§语言特征 仅为摘要。两者冲突时以 `voice_<领导>.md` 为准。

> ⚠️ 本文档中出现的 `知识库/方琳总/*.docx`、`assets/profile_方琳总_v1.yaml`、`outputs/2026Q3_讲话.docx` 等均为**示例路径**，实际执行时必须按上表规则解析为项目真实路径，**不得照搬**。

## 二、核心工作流（两阶段）

本 skill 严格按 **"先建模、后生成"** 的两阶段流程运行。两阶段可独立触发：

### 阶段 A：风格建模（首次或样本更新时执行）
```
样本目录 → extract_docx / extract_pptx / extract_pdf / extract_md
       → 解析格式 + 提取文本（输出同构的中间 JSON）
       → build_style_profile.py （多源分通道归并）
       → 风格档案 (YAML + Markdown 双格式)

语言维度（并行支线，v1.4 新增）：
样本目录 → dump_corpus.py（全量转文本）
       → pick_corpus.py（造通读选集）
       → mine_phrases.py（金句候选 + 四字词三桶）
       → 人工/模型通读复核
       → assets/voice_<领导>.md（语言手册） + 回填档案 linguistic.voice
```

> **多源分通道归并**（v3，2026-09-11）：
> - 把 docx / pdf / md 拆通道各自统计，再按字段权威性合成：`docx` 是格式权威源；`pdf` 作字号标题反推的补充证据；`md` 作文本/高频词补充
> - v3 新增 `format_hints` 通道：
>   - **md_frontmatter**：用户在 MD 顶部 YAML 显式声明的字体/字号字段（`font_body`/`size_body_pt`/`font_h1`/`line_spacing_pt`/`page_margin_*_mm`/`source_intent` 等）
>   - **inferred_from_leader**：通过 `--inferred-from-leader <领导>` 从 skill 内置 `profile_<leader>_v*.yaml` 注入的跨类型字体默认值（用于 MD 没有 frontmatter 但已知出自某位领导的情况）
> - 字段权威性矩阵（v3 扩展）：
>   ```
>   body_font           : docx > pdf > md_frontmatter > inferred_from_leader
>   body_size_pt        : docx > pdf > md_frontmatter > inferred_from_leader
>   line_spacing        : docx > md_frontmatter > inferred_from_leader  (pdf 无信号)
>   first_line_indent_cm: docx > md_frontmatter > inferred > pdf_bool
>   page_margins        : docx > md_frontmatter > inferred_from_leader  (pdf 无信号)
>   h1/h2/main_title 字体: docx > md_frontmatter > inferred_from_leader  (pdf 仅给 size_pt)
>   ```
> - 字段缺失 → WARNING + 在 `source_breakdown.<field>.selected_source` 标记来源，**绝不静默**
> - 非 docx 来源字段 → 顶层 `warnings` 提示"准确度有限"
> - 详见 `build_style_profile.py` 顶部说明 + 档案 `warnings` + `format.<field>_source` 字段

### 阶段 B：风格应用（每次有新报告需求时执行）
```
用户需求(主题/要点/大纲)
   → 双源加载：格式档案 profile_*.yaml  +  语言手册 voice_<领导>.md
   → 按文种选句式（O/B/C 系列）+ 按密度红线嵌金句 + 按词汇层遣词
   → 模型撰写 Markdown 源稿（generate_report.py 仅为骨架，勿依赖）
   → md_to_gongwen_docx.py --leader <领导> --doc-type <文种>  （落版）
   → 逐项自检（格式层 + 语言层） → 输出 .docx / .md / .pdf
```

## 三、阶段 A：风格建模详细步骤

### 1. 样本收集
- **样本来源**：知识库中 `领导姓名/` 文件夹下的全部历史文件 + 任意"规范.docx"。
- **最低样本数**：建议 ≥ 3 份（太少则风格不稳定），理想 ≥ 8 份。
- **样本类型平衡**：建议涵盖不同场景（致辞、汇报、总结、调研），保证风格档案覆盖度。

### 2. 格式提取（脚本：scripts/extract_*.py）

按文件类型分流提取以下要素：

| 维度 | docx | pptx | pdf | md |
|------|------|------|-----|-----|
| 字体（family） | ✅ run.font.name | ✅ run.font.name | ⚠️ **渲染字体归一**（按 FONT_ALIASES 归类） | ❌ 无 |
| 字号（size） | ✅ Pt 值 | ✅ Pt 值 | ✅ **字符级 size 字段 + 标题反推** | ❌ 无 |
| 颜色 | ✅ rgb | ✅ rgb | ❌ | ❌ 无 |
| 标题层级 | ✅ outlineLvl + style | ⚠️ 占位符类型 | ✅ **字号聚类反推** + page 锚定 | ✅ `#` 层级 + 分布统计 |
| 段落对齐 | ✅ alignment | ✅ alignment | ❌ | ⚠️ 文本推断 |
| 行距 / 段距 | ✅ Pt 值 | ⚠️ 部分支持 | ❌（line_spacing 记 None） | ❌ |
| 页边距 | ✅ sectPr | ⚠️ slideSize | ❌（page.margins 全 None） | ❌ 无 |
| 页码 / 页眉 | ✅ sectPr | ❌ | ⚠️ 文本识别 + y 坐标阈值过滤 | ❌ 无 |
| 强调/code | n/a | n/a | n/a | ✅ **手写解析器**，识别 `**`/`*`/`` ` `` |

→ 输出中间文件 `_intermediate/<文件名>.json`，结构示例：
```json
{
  "file": "方琳总_2025年度工作报告.docx",
  "format": {
    "main_title": {"font": "方正小标宋简体", "size_pt": 22, "color": "#000000"},
    "h1": {"font": "黑体", "size_pt": 16, "bold": false},
    "body": {"font": "仿宋_GB2312", "size_pt": 16, "line_spacing": "28pt", "first_line_indent": "2字符"},
    "page": {"top": "3.7cm", "bottom": "3.5cm", "left": "2.8cm", "right": "2.6cm"}
  },
  "text_samples": ["...", "..."]
}
```

### 3. 文本提取与语言分析
从每份样本中提取：
- 完整正文（去除页眉页脚、页码、图表说明）
- 段落首句集合（识别开篇模式）
- 段落末句集合（识别收尾模式）
- 高频四字词 / 短语
- 引用的政策表述 / 文件名 / 领导人讲话
- 称呼、署名、落款格式

### 4. 综合生成风格档案（脚本：build_style_profile.py）
脚本对全部中间 JSON 做**机械归并**（格式一致性统计、高频四字词、称呼/落款模式），产出档案骨架；**语义层的句式范式、结构范式、战略口号仍需模型通读样本后补齐**（脚本不调 LLM）。输出双格式档案：

**A. `assets/profile.yaml` — 机读版**（供 generate_report.py 使用）
```yaml
meta:
  leader: "方琳总"
  sample_count: 8
  generated_at: "2026-09-10"
format:
  fonts:
    main_title: {family: "方正小标宋简体", size_pt: 22, color: "000000"}
    h1: {family: "黑体", size_pt: 16}
    h2: {family: "楷体_GB2312", size_pt: 16}
    body: {family: "仿宋_GB2312", size_pt: 16}
  paragraph:
    line_spacing: 28          # 固定值（磅）
    first_line_indent: "2字符"
    space_before: 0
    space_after: 0
  page:
    margins: {top: 37, bottom: 35, left: 28, right: 26}  # mm
    page_number: "bottom-center"
linguistic:
  sentence_patterns:
    opener:
      - "在……的关键时期"
      - "值此……之际"
      - "本次大会的主要任务是"
    body:
      - "……，……，……"
      - "一是要……，二是要……，三是要……"
    closer:
      - "让我们更加紧密地团结在……周围"
      - "为谱写……新篇章而努力奋斗"
  vocabulary:
    high_freq_4char:
      - "踔厉奋发"
      - "勇毅前行"
      - "真抓实干"
      - "久久为功"
    political_terms:
      - "贯彻新发展理念"
      - "构建新发展格局"
      - "高质量发展"
  structure:
    typical_outline:
      - "一、回顾过去工作（成绩 + 不足）"
      - "二、分析当前形势（机遇 + 挑战）"
      - "三、部署下一步任务"
      - "四、号召与要求"
signature:
  format: "署名右对齐，单位+日期"
  example: "方  琳\n2026年9月10日"
```

**B. `assets/profile_<leader>.md` — 人读版**（供人工审阅、修订）
> 一份 Markdown 形式的风格说明，含具体示例与异常标注。

### 5. 风格档案的版本控制
- 每次重新建模生成新档案时，**保留历史档案**（`profile_<leader>_v1.yaml`、`v2.yaml` ...）。
- 命名规范：`profile_<领导名>_<文种>_v<版本号>.yaml`（文种=总结/致辞/呈报，如 `profile_方琳_总结_v1.yaml`）；生成时间写入档案内部的 `meta.generated_at`，**不放进文件名**，避免下游按名匹配失败。
- 随包内置：总结类 `profile_龚利民_总结_v1.yaml` / `profile_方琳_总结_v1.yaml`，文种专属 `profile_龚利民_致辞_v1.yaml` / `profile_龚利民_呈报_v1.yaml`。

## 四、阶段 B：风格应用详细步骤

### 1. 接收用户输入
明确向用户询问：
| 询问项 | 必填 | 说明 |
|--------|------|------|
| 报告类型 | ✅ | 致辞/汇报/总结/纪要/调研... |
| 主题 | ✅ | 例：2026 年三季度生产经营分析会讲话 |
| 大纲/要点 | ✅ | 用户提供，或由 skill 协助生成大纲后确认 |
| 字数范围 | ⚠️ | 默认 3000-5000 字 |
| 目标领导风格 | ⚠️ | 默认沿用已建模的领导；可临时切换 |
| 输出格式 | ⚠️ | 默认 .docx，可选 .md / .pdf |

### 2. 加载风格档案 + 语言手册（**双源加载，缺一不可**）

**2.1 加载格式档案**
- 文种专属档案优先：speech → `assets/profile_<leader>_致辞_v*.yaml`；baogao → `assets/profile_<leader>_呈报_v*.yaml`；否则 `assets/profile_<leader>_v*.yaml`。
- 若有多位领导档案，按用户指定切换；未指定时优先使用最新生成的一份。

**2.2 加载语言手册（🎙️ 阶段 B 的"音色"来源）** —— 按以下优先级：
1. `assets/voice_<领导>.md` ← **首选**，领导专属语言风格与金句手册
2. 档案 `linguistic.voice` 块 ← 机读兜底（含 `tone_profile` / `sentence_paradigms` / `golden_phrases` / `vocabulary_layers` / `forbidden`）
3. `references/linguistic_patterns.md` ← 通用兜底（无领导专属手册时才用）

> ⚠️ **格式与语言是两条线**：格式看 `profile_*.yaml` 的 `format` 段（实测值，不得推断）；语言看 `voice_<领导>.md`（语料挖掘，不得编造）。
> ⚠️ **不可跨领导借用金句** —— 龚董的「锐始者必图其终」不能出现在方琳稿里，反之方琳的「安全是'1'，没有安全，一切归'0'」也不能给龚董。串台一票否决。

### 3. 检索相似段落（可选增强）
- 对用户给出的要点做向量化（用本地 sentence-transformers，无需联网）。
- 在历史样本中检索 Top-5 相似段落，作为风格参考。
- 提升生成结果的"贴合度"。

### 4. 内容撰写（由模型完成，非脚本）
基于**格式档案 + 语言手册**，先产出 **Markdown 源稿**：

**4.1 格式四层（看 `profile_*.yaml`）**
- **段落级**严格遵循档案中的 `paragraph` 与 `fonts.body`（⚠️ 字号/缩进**以档案实测值为准**，龚利民为小二 18pt / 1.28cm，方琳为三号 16pt / 1.13cm，**不得写死"三号"**）。
- **结构级**严格遵循档案 `structure` 的文种范式（总结：四段式/12字；致辞：五段式；呈报：五段式）。
- **格式级**严格遵循：标题字体/字号、称呼/主送机关、落款格式、页边距。
- **落版级**：由 `md_to_gongwen_docx.py --leader <领导> --doc-type <文种>` 精确落地，见 §五.7。

**4.2 语言四层（看 `voice_<领导>.md`）** —— 这是"像不像那位领导"的关键：
- **语气层**：逐条比对手册 §一 `tone_profile`（含 ✅ 正例 / ❌ 反例），确保通读能听出该领导的"腔"。
- **句式层**：按文种从手册 §二 选用对应句式代号（开场 O系列 / 主体 B系列 / 收尾 C系列）。
- **金句层**：从手册 §三 按功能取用，**严格遵守密度红线**（总结 3–5 处 / 致辞 1–2 处 / 呈报 0–1 处），位置参考手册 §八。
- **词汇层**：优先手册 §四 的 L1 四字词、L2 战略/商业热词、L3 专名白名单（⚠️ 专名**不可简写、™不可省**）。
- **红线**：手册 §六 禁用清单 + §八 改写限度（引语原文/含数字表述/他人金句 ⛔ 不可改）。**无出处引语与编造数据是最高红线。**

> ⚠️ `generate_report.py` 是**骨架占位实现**，不调用 LLM，**不要依赖它写作**。真正的正文由模型撰写 Markdown 源稿，再由 `md_to_gongwen_docx.py` 落版。
> ⚠️ **金句只回答"怎么说话"，不提供事实**。一切数据 / 项目名 / 时间节点必须来自用户或知识库检索，缺失即标「待补充」——**手册中的金句是语料原话，不得当作新稿的事实来源**。

**文种分道（重要）——总结 / 致辞 / 呈报 三模板互斥，写什么像什么**：
- **report 总结类**（半年/年终总结、汇报）：四字对仗标题 + 黑体副标题 + 文末个人落款 + 页码
- **speech 致辞/讲话类**（致辞、讲话、主旨发言、就职演讲）：题注式，详见 `assets/profile_龚利民_致辞_v1.md`
- **baogao 呈报类**（对外汇报、专题汇报、经验交流材料）：公文报告体式，详见 `assets/profile_龚利民_呈报_v1.md`
- 致辞类五段式（实测自 GIIC 鸿蒙大会等 5 份样本）：
  1. 标题区：`# 在XXX上的致辞`（**不用四字对仗主标题**——那是总结专属）
  2. 题注：`## 深圳环水集团党委书记、董事长 龚利民`（署名行）+ `## （YYYY年M月D日）`（日期行）——**文末不落款**
  3. 称呼：`尊敬的XXX，各位领导、来宾、朋友们：`（对外场合点名来宾；内部场合可"同志们："）
  4. 主体：问候 → 场合意义 → 集团成效（**一是/二是/三是内嵌，不用黑体标题、不用（一）（二）条款行**）→ 倡议
  5. 结尾：对仗格言收束 + `谢谢大家！` + `**（共N字，约M分钟）**`（字数注记，缺失时落版脚本自动补算）
- 致辞篇幅 **800-1600 字（5-8 分钟）**，数据点到即止，不布置任务、不列责任单位。
- 呈报类五段式（实测自水利部报告、鸿蒙生态报告等 4 份样本）：
  1. 标题：`# （单位）关于XX（情况）的报告`——小标宋 22pt 居中，**"关于"句式**
  2. 主送机关/称呼：`水利部：`（部委报送）或 `尊敬的XX市长：`（领导专报）——**顶格**；经验材料式省略
  3. 开篇：单位一句话定位（"XX集团是XX企业，始终牢记……"）或呼应领导部署（"在您的亲自部署和推动下……"）
  4. 主体：`一、`黑体对仗标题 → `（一）`楷体条款 → 事实+数据展开，3-5 个一级标题（**与总结相同两级编号，但正文 16pt**）
  5. 结尾：`专此报告。` →（可选）`附件：XXX` → `**深圳市环境水务集团有限公司**` + `**YYYY年M月D日**`——**单位全称落款，无个人姓名**
- 呈报类正文 **16pt / 1.13cm**（区别于总结的 18pt！），语气客观恭敬、数据密集，无口号式动员、无"同志们"称呼。
- 落版时**必须**带对应 `--doc-type`：总结默认 report，致辞/讲话/就职演讲 `--doc-type speech`，对外/专题汇报 `--doc-type baogao`（中文别名均可），见 §五.7。

### 5. 输出文档
- 先输出一份 `.md` 源稿（可二次编辑、便于人工审阅）。
- 再用 `scripts/md_to_gongwen_docx.py --leader <领导>` 把 Markdown 转成 `.docx`，版式由脚本内置档案精确还原。
- 如需 PDF：先转 docx → 用 `docx2pdf` / `LibreOffice` 转 PDF。

### 6. 自检清单（生成后必须逐项检查）

**格式层**
- [ ] 标题字体是否为档案指定的主标题字体（党政公文通常"方正小标宋简体"二号）？
- [ ] 一级标题字体/字号是否与档案一致（通常"黑体"）？
- [ ] 正文字体/字号/行距/首行缩进是否**与风格档案实测值逐项对齐**（不是凭印象的"三号 + 2 字符"）？
- [ ] 页边距是否符合档案规范？
- [ ] 全文是否无英文标点、半角数字未规范处理？
- [ ] 落款格式（领导名/单位全称 + 日期）是否正确？
- [ ] 生成的 `.docx` 是否由 `md_to_gongwen_docx.py` 落版，且 `--leader` / `--doc-type` 参数正确？

**语言层（🎙️ 逐条对照 `voice_<领导>.md`）**
- [ ] 通读一遍，能听出该领导特有的"腔调"吗？还是像一份通用公文？
- [ ] 是否按文种选用了手册 §二 对应的句式代号（开场/主体/收尾三段）？
- [ ] **金句数量**是否在密度红线内（总结 3–5 / 致辞 1–2 / 呈报 0–1）？位置是否合理（破题 / 情绪高点 / 收尾）？
- [ ] 金句是否全部出自**该领导**手册 §三（**没有借用另一位领导的金句**）？
- [ ] 词汇是否落在该领导手册 §四 的层级内（L1 四字词密度、L2 术语、L3 专名完整未简写）？
- [ ] 是否触犯手册 §六 禁用清单（龚董：成语铺陈缺失/口号；方琳：无数据/无自省/引古文）？
- [ ] 引语原文、含数字表述、他人金句是否**一字未改**？（改写限度见手册 §八）
- [ ] 是否存在**无出处的引语或编造的数据**？（**最高红线**，有则立即改为「待补充」）
- [ ] 是否误入了另一领导的语域（人称 / 章节标题形式 / 字号缩进 / 标志性比喻）？

**致辞/讲话类（--doc-type speech）追加自检**：
- [ ] 主标题是否"在XXX上的致辞/讲话"句式（而非四字对仗）？
- [ ] 署名行（楷体 16pt 居中）+ 日期行（仿宋 16pt 居中）是否在标题下方（题注式）？
- [ ] 文末是否**无落款**、有"（共N字，约M分钟）"楷体右对齐注记？
- [ ] 正文是否无黑体标题层级、分点内嵌"一是/二是/三是"？
- [ ] 页脚是否**无页码**？
- [ ] 篇幅是否 800-1600 字？是否点到即止、未布置任务？

**呈报类（--doc-type baogao）追加自检**：
- [ ] 主标题是否"（单位）关于XX（情况）的报告"句式？
- [ ] 主送机关"XX部："或称呼"尊敬的XX市长："是否顶格？
- [ ] 正文是否 **16pt / 1.13cm**（勿沿用总结的 18pt）？
- [ ] "一、"是否黑体、"（一）"是否楷体？
- [ ] 是否有"专此报告。"结尾 + **单位全称落款（无个人姓名）** + 日期右对齐？
- [ ] 页码是否存在？
- [ ] 是否无"同志们"、无口号式动员、无致辞式格言收尾？
- [ ] 数据是否全部来自知识库/用户核实（呈报类数据责任最重）？

## 五、脚本使用方式（scripts/）

### 1. `extract_docx.py`
```bash
python scripts/extract_docx.py \
  --input "知识库/方琳总/*.docx" \
  --output "_intermediate/docx_samples.json"
```

### 2. `extract_pptx.py`
```bash
python scripts/extract_pptx.py \
  --input "知识库/方琳总/*.pptx" \
  --output "_intermediate/pptx_samples.json"
```

### 3. `extract_pdf.py`（v3：扫描件自动 OCR）
```bash
python scripts/extract_pdf.py \
  --input "知识库/<领导>/*.pdf" \
  --output "_intermediate/pdf_samples.json"

# 强制关闭 OCR（仅文本层）
python scripts/extract_pdf.py --input ... --output ... --no-ocr

# 提高 OCR 精度（默认 DPI 200，调到 300 慢但更准）
python scripts/extract_pdf.py --input ... --output ... --ocr-dpi 300
```
输出 JSON 结构（与 docx 同构 + meta）：
```json
{
  "file": "...pdf",
  "file_type": "pdf",
  "page_count": 30,
  "format_profile": {
    "fonts": {
      "body": {"family": "仿宋_GB2312", "size_pt": 16, "color": "000000",
               "text_source_mix": "text_layer"},     // v3 新增：text_layer / ocr / ocr+text_layer
      "fonts_dist": [["仿宋_GB2312", 230], ["黑体", 32]],
      "sizes_dist": [[16.0, 220], [18.0, 80]]
    },
    "paragraph": {
      "first_line_indent_inferred": true,
      "indent_distribution": {"indent": 80, "no_indent": 12},
      "line_spacing": null, "line_spacing_rule": null,
      "space_before_pt": null, "space_after_pt": null
    },
    "page": {"paper_width_mm": 210, "paper_height_mm": 297,
             "margin_top_mm": null, "margin_bottom_mm": null,
             "margin_left_mm": null, "margin_right_mm": null},
    "headings": [{..., "page": 1, "text_source": "text_layer"|"ocr"}]   // v3 新增
  },
  "text_profile": {
    "total_chars": 9200, "paragraph_count": 142,
    "paragraphs": ["...", "..."],
    "first_sentences": ["...", "..."],
    "last_sentences": ["...", "..."]
  },
  "warnings": [
    "PDF 首行缩进比例 < 50%，缩进判断不可靠",
    "PDF 页边距无可靠信号，已记 None"
  ],
  "meta": {                                         // v3 新增
    "ocr_available": true,
    "ocr_enabled": true,
    "ocr_pages": 0,
    "text_layer_pages": 30,
    "ocr_avg_confidence": null
  }
}
```

> **v3 扫描件 OCR**：
> - 检测：`page.extract_text()` 长度 < 50 字且 `chars < 5` → 视为扫描件
> - OCR：`rapidocr_onnxruntime`（跨平台 ONNX 模型，无需 tesseract.exe），支持中英
> - 输出：OCR bbox 坐标转 pdfplumber chars 格式（字号 = bbox 高度 × 0.75 估算），后续 y 坐标聚类零改动
> - 依赖：`pip install rapidocr-onnxruntime pillow`
> - 未安装包时 graceful 降级（`ocr_available=false`），仅跑文本层
> - 全文扫描件 PDF 会在 warnings 提示 `全部 N 页均为扫描件，已走 OCR（avg conf=X.XXX）`
>
> ⚠️ PDF 的 body_font 是渲染字体（已按 FONT_ALIASES 归类），**优先级低于 docx**——多源归并时若 docx 存在，body_font 取自 docx 而非 pdf。OCR 估算的字号精度弱于 docx 实测与 pdf 文本层。

### 4. `extract_md.py`（手写解析器，PyYAML + frontmatter + 跨类型字体推断，v2）
```bash
# 基本用法
python scripts/extract_md.py \
  --input "知识库/<领导>/" \
  --output "_intermediate/md_samples.json"

# MD 无 frontmatter 但确知出自某领导时：用 --inferred-from-leader 注入字体默认值
python scripts/extract_md.py \
  --input "知识库/方琳总/*.md" \
  --output "_intermediate/md_samples.json" \
  --inferred-from-leader "方琳"
```

**frontmatter 格式**（MD 顶部 YAML 块，v2 支持）：
```markdown
---
font_body: 仿宋_GB2312
size_body_pt: 18
font_h1: 黑体
size_h1_pt: 16
font_main_title: 方正小标宋简体
size_main_title_pt: 22
first_line_indent_cm: 1.28
line_spacing_pt: 28
page_margin_top_mm: 37
page_margin_bottom_mm: 35
page_margin_left_mm: 28
page_margin_right_mm: 26
source_intent: "推断自龚董 2026 半年总结样本"
---

# 这里是正文...
```

输出 JSON 结构（同构 + md 特有 blocks + format_hints）：
```json
{
  "file": "...md",
  "file_type": "md",
  "format_profile": {
    "fonts": {
      "body": {"family": "仿宋_GB2312", "size_pt": 18, "color": null, "_source": "md_frontmatter"},
      "h1":   {"family": "黑体", "size_pt": 16, ...},
      "main_title": {"family": "方正小标宋简体", "size_pt": 22, ...}
    },
    "paragraph": {"first_line_indent_cm": 1.28, "line_spacing": 28, ...},
    "page": {"margin_top_mm": 37, "margin_bottom_mm": 35, "margin_left_mm": 28, "margin_right_mm": 26, ...},
    "headings": [...],
    "headings_count": 26
  },
  "text_profile": {
    "total_chars": 5200, "paragraph_count": 39,
    "paragraphs": ["..."], "blocks": [...],
    "first_sentences": [...], "last_sentences": [...]
  },
  "block_stats": {"heading": 26, "paragraph": 31, "code": 13, "list": 3, "quote": 5, "divider": 10},
  "format_hints": {
    "md_frontmatter": {
      "applied_fields": {"font_body": "仿宋_GB2312", "size_body_pt": 18, ...},
      "raw": {"source_intent": "..."}
    }
  },
  "warnings": ["MD 格式字段来自 md_frontmatter（10 个）：font_body, size_body_pt, ..."],
  "meta": {"has_frontmatter": true, "frontmatter_keys": [...], "inferred_from_leader": null}
}
```

> **优先级**：frontmatter 字段 > inferred_from_leader 字段（同字段 frontmatter 优先）。
> 字段缺失仍走原逻辑（记 None + warning），且 `format_hints` 区分两个 source 通道。

### 5. `build_style_profile.py`（多源分通道归并，v3 + format_hints）
```bash
python scripts/build_style_profile.py \
  --intermediate "_intermediate/*.json" \
  --leader "方琳" \
  --output "assets/profile_方琳_总结_v1.yaml"
```
**核心机制**（v3 新增）：
- 按 `file_type` 把样本拆成 `docx / pdf / md` 三通道，各自 Counter
- 合成时按字段权威性：`docx > pdf > md_frontmatter > inferred_from_leader`
- v3 新增字段权威性矩阵（覆盖正文/标题/页边距/缩进/行距）
- 字段缺失则在 `source_breakdown.<field>.selected_source` 标记，且计入顶层 `warnings`，**绝不静默**
- 非 docx 来源字段 → `format.<field>_source` 标 `md_frontmatter` 或 `inferred_from_leader`
- 输出档案结构：
```yaml
meta:
  sample_breakdown: {docx: 3, pdf: 1, md: 1}    # v2 新增
  format_hints_usage:                            # v3 新增
    md_with_frontmatter: 1
    md_with_inferred: 0
  version: v3
format:
  body_font: 仿宋_GB2312                         # docx 优先
  body_font_source: docx                          # v3 新增
  body_size_pt: 16
  body_size_pt_source: docx                       # v3 新增
  line_spacing: 28
  line_spacing_source: docx                       # v3 新增
  first_line_indent_cm: 1.13
  first_line_indent_source: docx                  # v3 新增
  page_margins_mm: {top: 37, bottom: 35, left: 28, right: 26}
  page_margins_source: docx                       # v3 新增
  h1_font: 黑体
  h1_font_source: docx                            # v3 新增
  headings: [...]                                # docx + pdf 合并
  md_heading_level_distribution: {...}           # v2 新增
source_breakdown:                                # v3 扩展
  body_font:
    docx: {仿宋_GB2312: 230}                       # docx 分布
    pdf: {仿宋_GB2312: 80}                          # pdf 渲染字体
    md_frontmatter: 仿宋_GB2312                    # v3 新增
    inferred_from_leader: null                     # v3 新增
    selected_source: docx                          # v3 新增
    selected_value: 仿宋_GB2312
  body_size_pt:
    docx: {16.0: 220}
    pdf: {16.0: 60}
    md_frontmatter: 16.0                           # v3 新增
    inferred_from_leader: null                     # v3 新增
    selected_source: docx                          # v3 新增
    selected_value: 16.0
  ...
text_source_dist: {docx: 230, pdf: 142, md: 39}
warnings:
  - "body_font 来自 md_frontmatter（准确度有限，docx/pdf 缺失时启用）"   # v3 新增
  - "first_line_indent_cm 来自 inferred_from_leader（docx 缺失）"        # v3 新增
  ...
```

### 6. `generate_report.py`（骨架占位，勿依赖）
```bash
python scripts/generate_report.py \
  --profile "assets/profile_方琳_总结_v1.yaml" \
  --topic "2026年三季度生产经营分析会讲话" \
  --outline "1.成绩 2.形势 3.任务 4.号召" \
  --length 4000 \
  --output "outputs/2026Q3_讲话_方琳.docx"
```
> ⚠️ 本脚本**不调用 LLM**，仅演示如何读取档案并套用页面/字体设置。实际正文由模型撰写。

### 7. `md_to_gongwen_docx.py` ⭐ 阶段 B 的关键落版脚本
把 Markdown 报告转成集团公文档式 docx，内置两位领导的实测版式档案，**无硬编码路径，跨项目可直接使用**。
```bash
# 总结/汇报类（默认 --doc-type report）
python scripts/md_to_gongwen_docx.py \
  --input "outputs/2026Q3_讲话.md" \
  --output "outputs/2026Q3_讲话_方琳.docx" \
  --leader "方琳"

# 致辞/讲话/就职演讲类（题注式，必须带 --doc-type speech）
python scripts/md_to_gongwen_docx.py \
  --input "outputs/水务科技AI分享会致辞.md" \
  --output "outputs/20260911_会议致辞_..._龚利民.docx" \
  --leader "龚利民" --doc-type speech

# 对外汇报/专题汇报/经验材料（公文报告体式，--doc-type baogao）
python scripts/md_to_gongwen_docx.py \
  --input "outputs/关于XX工作情况的报告.md" \
  --output "outputs/20260911_对外汇报_..._龚利民.docx" \
  --leader "龚利民" --doc-type baogao
```
`--doc-type` 接受 `report`/`speech`/`baogao` 及中文别名（致辞/讲话/发言/总结/报告/汇报/呈报/对外汇报/专题汇报/经验材料等）。speech 优先加载 `assets/profile_<领导>_致辞_v*.yaml`，baogao 优先加载 `assets/profile_<领导>_呈报_v*.yaml`，无则回退总结档案 + 文种渲染规则。

内置版式（`--leader` 支持中文名与别名：龚董/董事长、方总/总裁）：

| 领导 | 正文字号 | 首行缩进 | 页边距(上/下/左/右 mm) |
|------|---------|---------|---------------------|
| 龚利民 | 仿宋_GB2312 小二 18pt | 1.28cm | 37 / 35 / 26 / 26 |
| 方琳 | 仿宋_GB2312 三号 16pt | 1.13cm | 37 / 35 / 28 / 26 |
| 未识别 | 回退三号 16pt | 1.13cm | 37 / 35 / 28 / 26 |
| 龚利民·致辞(`--doc-type speech`) | 仿宋_GB2312 小二 18pt | **1.27cm** | 37 / 35 / **28** / 26 |
| 龚利民·呈报(`--doc-type baogao`) | 仿宋_GB2312 **三号 16pt** | **1.13cm** | 37 / 35 / **28** / 26 |

共同项：主标题 方正小标宋简体 22pt 居中；副标题 黑体 18pt 居中；`（一）`类条款用楷体；行距固定 28pt。**致辞类无页码、无副标题、文末不落款。**

**Markdown → docx 映射约定（report 模式）**：

| Markdown | docx 效果 |
|----------|----------|
| `# 主标题` | 居中 方正小标宋简体 22pt |
| `## 副标题` | 居中 黑体 18pt |
| `### 一级标题` | 黑体，与正文同级字号 |
| `（一）`/`(一)` 开头 | 楷体条款段 |
| `XX：` 结尾且 ≤25 字 | 称呼语，**顶格不缩进**（如"同志们："） |
| `**整行**` | 落款，右对齐 |
| `> ...` | 引用块，**仅 md 保留，不写入 docx** |
| `---` / 空行 | 跳过 |

**Markdown → docx 映射约定（speech 模式，题注式）**：

| Markdown | docx 效果 |
|----------|----------|
| `# 在XXX上的致辞` | 居中 方正小标宋简体 22pt |
| `## 职务+姓名` | **署名行**：居中 楷体 16pt（题注式） |
| `## （YYYY年M月D日）` | **日期行**：居中 仿宋 16pt |
| `尊敬的……：` / `同志们：` | 称呼语，顶格不缩进 |
| 正文 | 仿宋 18pt、缩进 1.27cm；分点内嵌"一是/二是/三是" |
| `**（共N字，约M分钟）**` | 字数注记：楷体 16pt 右对齐（缺失自动补算，约 200 字/分钟） |

**Markdown → docx 映射约定（baogao 模式，公文报告体式）**：

| Markdown | docx 效果 |
|----------|----------|
| `# （单位）关于XX（情况）的报告` | 居中 方正小标宋简体 22pt |
| `水利部：` / `尊敬的XX市长：` | **主送机关/称呼**：顶格不缩进（自动识别） |
| 正文 | 仿宋 **16pt**、缩进 1.13cm、两端对齐 |
| `一、` 开头独立短行（≤50字、无句号） | **黑体 16pt** 一级标题（baogao 模式自动识别） |
| `（一）` 开头 | 楷体 16pt 条款段 |
| `专此报告。` / `附件：XXX` | 普通正文段（缩进同正文） |
| `**单位全称**` + `**YYYY年M月D日**` | 落款：右对齐 仿宋 16pt（**单位落款，非个人姓名**） |

> 📌 首行缩进与字号绑定：公文"首行缩进 2 字符"≈ 2 × 正文字号。18pt 对应 1.28cm，16pt 对应 1.13cm。换领导换字号时缩进必须同步换，不能沿用。

### 8. `dump_corpus.py` — 语料抽取（语言建模第一步）

把知识库里一位领导的**全部**样本转成纯文本语料，供金句挖掘与人工通读。
```bash
python scripts/dump_corpus.py \
  --kb-root "<知识库根目录>" \
  --out-dir "<工作目录>/corpus"
```
- 支持 `.docx` / `.pdf`（文本层）/ `.doc`（Word COM 转换）/ `.pptx`（文本框）
- 按一级目录自动分桶：`corpus_<领导>.txt`（含龚利民 / 方琳），其余归 `corpus_通用.txt`
- 产出 `dump_report.json` 记录每份文件状态（`ok` / `skip` / `warn` + 字数）
- ⚠️ 扫描件 PDF 文本层为空 → 报 `warn` 并跳过，需先走 `extract_pdf.py`（OCR）另处理

### 9. `mine_phrases.py` — 金句候选挖掘（语言建模第二步）⭐
```bash
# 金句 + 四字词（推荐起步参数：min-score 5 / top 400）
python scripts/mine_phrases.py \
  --corpus "corpus/corpus_龚利民.txt" \
  --out    "phrase_龚利民" \
  --min-score 5 --top 400

# 只出四字词
python scripts/mine_phrases.py --corpus ... --out ... --only-4char
```
- **启发式打分**（不调 LLM）：格言反差 / 号召动员 / 对仗 / 排比 / 比喻 / 成语密度 / 引语包装 / 递进转折；对数据堆砌、编号段落、事务表述罚分
- **四字词三桶**：`idioms`（真公文成语，可直接进 `high_freq_4char`）/ `domain_terms`（业务专名）/ `fragments_dropped`（长专名切片，如「千家万户水」← 「千家万户水管家」，靠扩展比剔除，避免污染词表）
- 产出 `<out>.json`（结构化，含 score / tags / source）+ `<out>.md`（按标签分组，便于逐条圈选）
- ⚠️ **脚本输出只是候选**，最终金句必须人工/模型复核后写入 `assets/voice_<领导>.md`，**不得把脚本输出直接当作金句库**（避免把专名切片或事务句当金句）

### 10. `pick_corpus.py` — 造通读选集
从语料里挑出指定文档拼成一份便于通读的选集（按关键词匹配文件名）。
```bash
python scripts/pick_corpus.py corpus/corpus_龚利民.txt 选集_龚董_致辞讲话.txt 致辞 讲话 论坛 大会
```

> **语言建模完整链路**：
> `dump_corpus.py`（全量转文本）→ `pick_corpus.py`（挑选集）→ `mine_phrases.py`（出候选）→ **人工/模型通读复核** → 撰写 `assets/voice_<领导>.md` + 回填 `profile_*.yaml → linguistic.voice`。

### 11. `sync_copies.py` — 跨副本一致性巡检与同步 ⚙️

本技能同时存在于插件包 / 项目级 skill / 专家分发包三处，插件外壳另有 P / R 两处。
**改动任一 `SKILL.md` / `assets/` / `scripts/` / 插件外壳文件后，必须跑一次**：

```bash
python scripts/sync_copies.py --check     # 巡检；有漂移或行尾腐坏 → 退出码 1，可接 CI
python scripts/sync_copies.py --apply     # 同步：skill 层 3 副本 + 插件层 2 副本
```

包含跨副本版本分叉台账（`TAKE_FROM` / `PLUGIN_TAKE_FROM`）、行尾腐坏（`\r\r\n`）体检与自愈。
**机制、三条铁律与最小验收流程见 §十** —— 不读 §十 就改文件，极易被旧版反向覆盖或引发表格落版静默失效。

## 六、参考资源（references/ 与 assets/）

### 6.1 语言资源加载优先级（🎙️ 阶段 B 必读顺序）

| 优先级 | 资源 | 何时加载 |
|:---:|------|---------|
| **1** | `assets/voice_<领导>.md` | **首选**。只要目标领导有专属语言手册，就必须加载（`龚利民` / `方琳` 均已具备） |
| **2** | 档案 `linguistic.voice` 块 | 机读兜底：`profile_*.yaml` 中的 `tone_profile` / `sentence_paradigms` / `golden_phrases` / `vocabulary_layers` / `forbidden` |
| **3** | `references/linguistic_patterns.md` | **通用兜底库 + 领导音色路由**——无专属手册，或手册未覆盖的文种时加载 |

> ⚠️ 三者**不叠加使用**：`voice_<领导>.md` 已内含该领导的语气、句式、金句、词汇、禁用五层，直接照它写即可；只有在手册缺失时才降到下一级。
> ⚠️ `linguistic_patterns.md` 是**通用**模板库（含"领导音色速查对照表"与"串台防范三条铁律"），它不会告诉你"龚董会怎么说"，只会告诉你"公文一般怎么说"。

### 6.2 其他参考资源

| 资源 | 何时加载 |
|------|---------|
| `references/format_standards.md` | 当用户要求"严格按党政公文标准"或档案中检测到疑似党政格式时必读 |
| `references/style_dimensions.md` | 当风格提取需要扩展新维度（如多语言、表格、图片）时读取 |

不要把所有 references 内容堆在 SKILL.md 中，按需加载以节省上下文。

## 七、关键约束与质量红线

1. **绝不联网**：所有解析、生成都在本地完成，避免敏感领导讲话外泄。
2. **绝不复用其他领导档案**：每位领导单独建档，避免风格串台（**金句不可跨领导借用**，一票否决）。
3. **保留人工审阅环节**：skill 输出仅作"初稿"，必须由综合办公室人员人工复核、签字。
4. **敏感词过滤**：避免生成包含绝对化表述、夸张修辞、政治不正确的文本（即使是领导文风，也要确保合规）。
5. **不可伪造署名**：领导姓名、职务、日期必须由用户确认，skill 不擅自填写。
6. **🎙️ 金句只提供"怎么说话"，不提供事实**：`voice_<领导>.md` 中的金句是**语料原话**，可用于模仿语气与句式；但一切**数据 / 项目名 / 时间节点 / 政策依据**必须来自用户提供或知识库检索，缺失即标「待补充」，**严禁由手册或模型编造**。无出处引语与编造数据是**最高红线**。
7. **🎙️ 金句密度封顶**：总结 3–5 处 / 致辞 1–2 处 / 呈报 0–1 处。超密度 = 通篇口号，视为质量不合格。
8. **🎙️ 改写限度**：句式结构（对仗/比喻/自省句式）可套用；**引语原文、含数字表述、他人金句 ⛔ 不可改写**（详见各手册 §八）。

## 八、典型用例

| 用户输入 | skill 行为 |
|---------|-----------|
| "分析一下方琳总的报告风格" | 触发阶段 A，从样本目录抽取样本，生成 `profile_方琳_总结_v1.yaml` + `voice_方琳.md` |
| "按方琳总的风格写一篇 2026 年三季度讲话" | 触发阶段 B，加载 `profile_方琳_总结_v1.yaml` + `voice_方琳.md` → 生成 .docx |
| "龚董和方琳总的风格有什么区别？" | 同时加载两份档案 + 两份语言手册，做 diff 输出对比表 |
| "帮我把领导的金句库再丰富一些" | 跑 `dump_corpus.py` → `mine_phrases.py` → 通读复核 → 扩充 `voice_<领导>.md` 与档案 `linguistic.voice` |
| "这句话是龚董说过的吗？" | 在 `voice_龚利民.md` §三 逐条比对，**标了出处的才是原话，未标的是句式模板** |
| "把这份材料改成符合公文标准的格式" | 调用 `format_standards.md` + 现有档案，校正版式 |

## 九、扩展点（后续可迭代）

- ✅ ~~增加 OCR 模块（扫描件 PDF 通过 OCR 提取文本）~~ — v3 已实现（rapidocr-onnxruntime）
- ✅ ~~按文种（致辞/总结）分道模板~~ — 已实现（`--doc-type speech` + `profile_<领导>_致辞_v*.yaml`，2026-09-11）
- ✅ ~~呈报类模板（对外汇报/专题汇报/经验材料）~~ — 已实现（`--doc-type baogao` + `profile_龚利民_呈报_v1.yaml`，2026-09-11）
- ✅ ~~语言维度扩充（语气 / 金句 / 句式 / 词汇）~~ — 已实现（**v1.6**，2026-09-14）：`voice_龚利民.md`（18 份/12.0 万字/33 条金句）+ `voice_方琳.md`（3 份/2.2 万字/31 条金句）；四份档案均追加 `linguistic.voice` 机读块
- ✅ ~~4 字词过滤白名单~~ — 已实现（`mine_phrases.py` 的**四字词三桶**：`idioms` / `domain_terms` / `fragments_dropped`，用"扩展比"剔除长专名切片，见 §五.9）
- 增加表格自动生成（领导报告中常含"主要指标完成情况"等表格）。
- 增加引语智能识别（"习近平总书记强调……"自动套用规范格式）。
- **补充方琳语料**（`voice_方琳.md` §九 已标注短板）：优先季度经营分析会讲话、专题部署讲话、对外致辞、对外汇报材料 —— 现有 3 份几乎全是半年度/年度总结，**缺致辞与呈报语体样本**，两类文种按方琳风格生成前须与用户确认语气基线。
- 方琳致辞/呈报档案（`profile_方琳_致辞/呈报_v*.yaml`）：待上述语料到位后按 §三 同法建档。
- 增加多领导档案融合（生成"集团公司文件"时融合多方文风）。
- 金句**时效性复核机制**：金句标注了文种与出处，但未标年份衰减——可增加"过时口号"标记（如已停止使用的战略表述）。

## 十、跨副本一致性维护（改文件后必读）

本插件的两个技能各自存在于**三处**，插件外壳（`plugin.json` / `README.md` /
`USER_GUIDE.md` / `agents/` / `avatars/`）又存在于**两处**：

| 层 | 副本 | 内容 |
|----|------|------|
| **skill 层** | **P** 插件包 / **Q** 项目级 skill / **R** 专家分发包 | 各技能的 `SKILL.md`、`assets/`、`scripts/`、`references/`、`selftest/` |
| **插件层** | **P** / **R** | 插件外壳（Q 是裸 skill 目录，无外壳） |

`sync_copies.py` 按 **`SKILLS` 注册表**逐技能巡检，当前管辖：
`zhibi-suiwen`（执笔随文）与 `knowledge-base-organizer`（找得到）。

> ⚠️ **新增技能必须登记进 `SKILLS` 注册表**，否则它的三副本漂移永远不会被发现。
> 2026-09-14 实际踩过：`knowledge-base-organizer` 漂移了 4 个文件，直到人工统一
> "总裁"称谓时才被翻出来 —— 它此前不在工具管辖范围内。

改动任意一处后，**必须**跑：

```bash
# 巡检（有漂移/行尾腐坏 → 退出码 1，可直接接 CI）
python scripts/sync_copies.py --check

# 同步：全部注册技能的 3 副本 + 插件层 2 副本
python scripts/sync_copies.py --apply

# 只处理某个技能 / 只处理 skill 层
python scripts/sync_copies.py --apply --skill zhibi-suiwen
python scripts/sync_copies.py --apply --no-plugin
```

### 三条铁律

1. **规范源默认是 P**。在 Q / R 侧修了 bug 却忘记登记，会被 P 侧旧版**反向覆盖**。
   例外文件必须登记进对应技能的 `TAKE_FROM`（skill 层）或 `PLUGIN_TAKE_FROM`（插件层）——
   这些表是跨副本版本分叉的**唯一台账**，也是复盘"谁覆盖了谁"的唯一依据。
2. **字节级 I/O**：脚本只用 `read_bytes` / `write_bytes`，绝不用 `read_text` / `write_text`。
   后者在 Windows 上会把 `\n` 翻译成 `\r\n`；若原文已是 CRLF，就写出 `\r\r\n` ——
   `str.splitlines()` 把这种行尾当成**两个**换行，Markdown 表格行之间因此多出空行，
   `md_to_gongwen_docx.py` 的 `parse_md_table` 判定「表头下一行不是分隔行」而**整块
   退化为普通段落**：落版静默失效，肉眼极难发现。2026-09-14 实际踩过一次，15 个文件中招。
   回归判据是 `selftest` 的 `C04` / `C07` 用例（两者都断言表格格数）。
3. **身份中性措辞**：`sync_copies.py` 会按目标副本改写技能名（插件包与分发包共用
   一个名字，项目级 skill 用另一个名字，映射写在脚本的 `NAME_P` / `NAME_Q` 常量里）。
   因此正文里**不要硬写具体技能名** —— 写了就会触发"两名并存 → 跳过替换"护栏，
   使 `frontmatter.name` 与正文标题无法按副本改写。若某文件**天生必须同时写出两个
   名字**（目前只有 `selftest/README.md` 与脚本自身），须登记进该技能的
   `no_transform`；其余新增内容一律用「本技能」这类中性措辞。
4. **改称/正名要留兼容轨**：统一称谓（如「方琳（执行总裁）」→「方琳（总裁）」）时，
   旧名**不要从代码里删干净** —— 归档目录别名、迁移映射里保留旧名是有意为之
   （已归档的历史素材仍用旧名）。2026-09-14 的改称里，`knowledge-base-organizer`
   的 `taxonomy.yaml` aliases、`backtest.py` 旧→新映射、`kbsync.py` 远端候选列表
   均保留 `02-方琳（执行总裁）`，**不要"顺手清理"**。

### 改动后的最小验收

```bash
python scripts/sync_copies.py --check          # 期望：退出码 0，无漂移、无行尾腐坏
cd selftest && python regress.py --out-dir out_v2 --json v2_verify.json
                                               # 期望：10 份 / 失败断言 0 项
```

> `selftest/out_*/` 与 `selftest/*verify*.json` 是运行产物，不纳入同步、不需提交；
> 验收完成后可随手删除。