#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mine_phrases.py —— 从语料中挖掘「金句候选」与「高频四字词」，用于风格建模的语言维度扩充。

设计目标：把"人工通读几十万字语料"降级为"人工复核几百条候选"。
本脚本只做**启发式打分与聚合**，不做语义判断；最终金句必须由人工/模型复核后写入
`assets/voice_<领导>.md`，**不得把脚本输出直接当作金句库**。

用法:
    python scripts/mine_phrases.py \
        --corpus "_backtest/corpus/corpus_龚利民.txt" \
        --out    "_backtest/phrase_candidates_龚利民" \
        --top 220

    # 只出高频四字词
    python scripts/mine_phrases.py --corpus ... --out ... --only-4char

产出:
    <out>.json  结构化：{candidates:[{text,score,tags,source}],
                        idioms:[...] 真四字词（命中内置种子表）,
                        domain_terms:[...] 高频业务术语（非成语）,
                        fragments_dropped:[...] 被剔除的长词碎片}
    <out>.md    人读版：按标签分组、带分数与出处，便于逐条圈选

    四字词三桶说明：
      · idioms          —— 命中 IDIOM_SEED（可按 --idiom-list 覆盖）的公文成语，可直接进档案 high_freq_4char
      · domain_terms    —— 高频但属业务专名/术语（科技创新、水质净化…），归 strategic/业务语汇
      · fragments_dropped —— 「千家万户水」「质净化厂」这类长专名的切片，靠"扩展比"剔除，
                           避免污染 high_freq_4char（修复 SKILL.md §九 挂账问题）

打分维度（命中即加分，权重见 WEIGHTS）:
    格言反差  不是…而是… / 没有…就…        决策句、判断句，领导金句最高发形态
    号召动员  让我们 / 必须 / 一定要 / 我们要
    对仗结构  分号或逗号切出的等长并列小句
    排比结构  ≥3 个小句同字起头
    比喻意象  如/像/好比/像…一样 / "1"…"0" / 压舱石 / 桥头堡 …
    成语密度  4 字词 ≥2 个（用语料内 n-gram 结果交叉验证）
    引语包装  中文引号包裹的短语
    递进转折  但 / 然而 / 更要看到 / 同时也要
    篇幅适中  12-52 字（太短无信息量、太长非金句）
罚分:
    数据堆砌  ≥4 个数字 / 出现 % 或 亿元 等度量
    编号段落  行首为 一、/（一）/ 1.
    事务表述  含"责任单位""牵头""台账""报送""附件"等
"""
import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

# ---------------- 配置 ----------------

WEIGHTS = {
    "格言反差": 6,
    "号召动员": 5,
    "对仗结构": 4,
    "排比结构": 4,
    "比喻意象": 5,
    "成语密度": 3,
    "引语包装": 2,
    "递进转折": 2,
    "篇幅适中": 2,
}

PENALTY = {
    "数据堆砌": -5,
    "编号段落": -6,
    "事务表述": -6,
}

PAT_GEM = re.compile(r"不是.{0,20}而是|没有.{0,12}就|不在于.{0,20}而在于|越是.{0,16}越")
PAT_CALL = re.compile(r"让我们|必须|一定要|我们要|务必|决不能|绝不能|应当|需要清醒")
PAT_METAPHOR = re.compile(
    r"好比|如同|犹如|像.{1,8}一样|「1」|“1”|压舱石|桥头堡|生命线|定盘星|总开关|"
    r"牛鼻子|最后一公里|最后一米|组合拳|指挥棒|风向标|晴雨表|先手棋|攻坚战|新引擎|硬骨头"
)
PAT_QUOTE = re.compile(r"[“「][^”」]{2,24}[”」]")
PAT_CONTRAST = re.compile(r"但|然而|更要看到|同时也要|另一方面|不能只看")
PAT_TXN = re.compile(r"责任单位|牵头|台账|报送|附件|报送材料|销号|挂钩考核")
PAT_NUM = re.compile(r"\d")
PAT_METRIC = re.compile(r"%|％|亿元|万元|万吨|万立方米|个百分点|万人|座|公里")
PAT_NUMBERED = re.compile(r"^\s*(?:[一二三四五六七八九十]+、|[（(][一二三四五六七八九十\d]+[）)]|\d+[.、])")
IDIOM_STOP = set("的了在是和与及或对为以及等这那有不也就都还很更最一二三四五六七八九十"
                 "上下前后左右中你我他她它此其孰么什怎样谁知")

# 公文常用四字词种子表（用于把真成语与业务专名区分开）。
# 随包内置，可用 --idiom-list 覆盖；未命中的 4 字词仍会经"扩展比"过滤后进入 domain_terms。
IDIOM_SEED = set("""
迎难而上 真抓实干 善作善成 苦干实干 敢打硬仗 能打胜仗 立柱架梁 砥砺奋进 凝心聚力 守正创新
踔厉奋发 勇毅前行 久久为功 埋头苦干 驰而不息 锲而不舍 求真务实 严谨细致 精益求精
攻坚克难 稳中有进 提质增效 降本增效 统筹推进 压实责任 标杆示范 稳中求进 转型升级
敢闯敢试 敢为人先 勇立潮头 逢山开路 遇水架桥 蹄疾步稳 有序推进 重点突破
多措并举 精准施策 靶向发力 攻坚发力 协同发力 系统发力 纵深推进 纵深发展
大胆探索 先行先试 示范引领 典型引路 以点带面 全面铺开 落地见效 见行见效
锚定目标 对标一流 争先进位 比学赶超 你追我赶 加压奋进 乘势而上
顺势而为 因势利导 危中寻机 化危为机 抢抓机遇 主动求变 主动作为
担当作为 履职尽责 尽心尽力 全力以赴 全神贯注 心无旁骛 一以贯之 常抓不懈
持之以恒 抓铁有痕 踏石留印 干在实处 走在前列 勇挑重担 令行禁止 雷厉风行
注重实效 务求实效 讲求实效 提质扩面 精细管理 科学管理 规范管理
防范风险 守住底线 筑牢防线 固本强基 夯基固本 强基固本 补齐短板 扬长补短
开源节流 增收节支 精打细算 颗粒归仓 应纳尽纳 应收尽收
远近结合 内外兼修 上下联动 条块结合 上下同心 同向发力 同频共振 协同联动
一抓到底 一鼓作气 趁热打铁 开拓进取 锐意进取 奋发有为 大有可为 开创新局
保底攀高 争先创优 争创一流 勇争一流 敢争第一 破立并举 先立后破 破旧立新
迭代升级 更新换代 龙头带动 做强做优 做大做强 做实做细 做深做透
锐始者必图其终 不进则退 慢进亦退 差之千里 时不我待 只争朝夕 只争朝タ
弯道超车 换道超车 变道超车 逆势而上 逆势增长 稳中有升 稳中向好 稳中提质
善作有为 抓细抓实 落地落细 见实见效 行稳致远 蹄疾步稳 笃行不怠
千帆竞发 百舸争流 奋楫者先 勇进者胜 不进则退
立说立行 即知即改 边学边改 真改实改 真学真信 学深悟透 入脑入心 走深走实
举一反三 以案促改 以案促治 精准运用 抓早抓小 防微杜渐 标本兼治
内外并举 软硬兼施 疏堵结合 建管并重 量质并举 质效并举
爬坡过坎 滚石上山 负重前行 逆水行舟 中流击水 稳扎稳打
干字当头 实干为要 实绩说话 以实为要 以干为要 不比表态 只看实绩
革故鼎新 推陈出新 吐故纳新 应变求变 识变求变 求新求变
""".split())


def normalize(text: str) -> str:
    t = text.strip()
    t = re.sub(r"\s+", "", t)
    return t


def split_sentences(text: str):
    """按句末标点切句，保留分号作为句内结构。"""
    raw = re.split(r"(?<=[。！？!?])", text)
    out = []
    for s in raw:
        s = normalize(s)
        if s:
            out.append(s)
    return out


def ngram_4char(text: str) -> Counter:
    """语料内四字词 n-gram 统计（含中文标点/字母的窗口丢弃）。"""
    cnt = Counter()
    cjk = re.compile(r"[\u4e00-\u9fff]")
    for i in range(len(text) - 3):
        w = text[i:i + 4]
        if not all(cjk.match(c) for c in w):
            continue
        if any(c in IDIOM_STOP for c in w):
            continue
        cnt[w] += 1
    return cnt


def ngram_k(text: str, k: int) -> Counter:
    """任意窗口长度的中文 n-gram 统计，用于「扩展比」碎片检测。"""
    cnt = Counter()
    cjk = re.compile(r"[\u4e00-\u9fff]")
    for i in range(len(text) - k + 1):
        w = text[i:i + k]
        if not all(cjk.match(c) for c in w):
            continue
        cnt[w] += 1
    return cnt


def split_fragments(four_grams: Counter, five_grams: Counter,
                    ratio: float = 0.55, min_freq: int = 3, min_ext_abs: int = 4):
    """把 4 字词拆成「真四字词」与「长词碎片」。

    碎片判定：若「该 4 字词 + 前后各补一字」形成的 5 字窗口，其最高频次同时满足
      · ≥ 该 4 字词频次 × ratio（相对证据）
      · ≥ min_ext_abs（绝对证据）
    则判为某个更长专名/术语的切片。
      例：「千家万户」50 次 →「千家万户水」48 次，相对 0.96 / 绝对 48 → 碎片
          「真抓实干」8 次 → 最高 5 字窗口「真抓实干，」不构成纯中文窗口 → 保留

    ⚠️ 绝对门槛不可省：低频 4 字词（含大量一次性窗口）的扩展窗口必然同样低频，
    仅看比值会把所有低频词误杀（实测曾一次剔除 22763 条）。
    """
    ext_after = Counter()   # 4字词 -> 其作为前缀时后续 5 字窗口的最高频
    ext_before = Counter()  # 4字词 -> 其作为后缀时前置 5 字窗口的最高频
    for g5, m in five_grams.items():
        ext_after[g5[:4]] = max(ext_after[g5[:4]], m)
        ext_before[g5[1:]] = max(ext_before[g5[1:]], m)

    keep, frags = [], []
    for w, n in four_grams.items():
        if n < min_freq:
            continue
        best = max(ext_after.get(w, 0), ext_before.get(w, 0))
        if best >= max(n * ratio, min_ext_abs):
            frags.append({"word": w, "count": n, "ext_count": best})
        else:
            keep.append((w, n))
    return keep, frags


def load_corpus(path: Path):
    """返回 [(source_label, full_text), ...]，兼容 dump_corpus.py 的 "#=== file: ... ===" 分隔。"""
    raw = path.read_text(encoding="utf-8", errors="replace")
    chunks = re.split(r"^#=== file: (.+?) ===$", raw, flags=re.M)
    docs = []
    if len(chunks) > 1:
        # chunks = [preamble, label1, body1, label2, body2, ...]
        for i in range(1, len(chunks) - 1, 2):
            label = chunks[i].strip()
            body = chunks[i + 1]
            docs.append((label, body))
    else:
        docs.append((path.name, raw))
    return docs


def score_sentence(s: str, idiom_hits: int):
    tags = []
    score = 0

    body = s.rstrip("。！？!?")
    n = len(body)

    if PAT_GEM.search(s):
        tags.append("格言反差")
    if PAT_CALL.search(s):
        tags.append("号召动员")
    if PAT_METAPHOR.search(s):
        tags.append("比喻意象")
    if PAT_QUOTE.search(s):
        tags.append("引语包装")
    if PAT_CONTRAST.search(s):
        tags.append("递进转折")

    # 对仗：按分号/逗号切小句，若存在 ≥2 个长度相近（差≤2）的小句
    clauses = [c for c in re.split(r"[，,；;]", body) if c]
    if len(clauses) >= 2:
        lens = [len(c) for c in clauses]
        for i in range(len(lens) - 1):
            if abs(lens[i] - lens[i + 1]) <= 2 and lens[i] >= 3:
                tags.append("对仗结构")
                break
    # 排比：≥3 个小句同字起头，或"一是/二是/三是"式
    if len(clauses) >= 3:
        heads = [c[0] for c in clauses]
        if len(set(heads)) <= max(1, len(heads) - 2) and len(heads) >= 3:
            tags.append("排比结构")
        elif re.search(r"一是.*二是.*三是", body):
            tags.append("排比结构")

    if idiom_hits >= 2:
        tags.append("成语密度")

    if 12 <= n <= 52:
        tags.append("篇幅适中")

    for t in tags:
        score += WEIGHTS.get(t, 0)

    # 罚分
    digit_runs = len(PAT_NUM.findall(body))
    if digit_runs >= 4 or (PAT_METRIC.search(body) and digit_runs >= 2):
        tags.append("数据堆砌")
        score += PENALTY["数据堆砌"]
    if PAT_NUMBERED.match(s):
        tags.append("编号段落")
        score += PENALTY["编号段落"]
    if PAT_TXN.search(s):
        tags.append("事务表述")
        score += PENALTY["事务表述"]

    return score, tags


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", required=True, nargs="+", help="一个或多个语料 txt")
    ap.add_argument("--out", required=True, help="输出文件前缀（不含扩展名）")
    ap.add_argument("--top", type=int, default=200, help="金句候选条数")
    ap.add_argument("--min-score", type=int, default=6, help="最低入选分")
    ap.add_argument("--top-4char", type=int, default=60, help="高频四字词条数")
    ap.add_argument("--min-4char-freq", type=int, default=3)
    ap.add_argument("--frag-ratio", type=float, default=0.55,
                    help="长词碎片判定阈值：5 字窗口频次/4 字词频次 ≥ 该值时判为碎片")
    ap.add_argument("--idiom-list", help="自定义四字词种子表（每行一个或空格分隔）")
    ap.add_argument("--only-4char", action="store_true")
    args = ap.parse_args()

    idiom_seed = IDIOM_SEED
    if args.idiom_list:
        idiom_seed = set(Path(args.idiom_list).read_text(encoding="utf-8").split())
        print(f"加载自定义四字词种子表：{len(idiom_seed)} 条", file=sys.stderr)

    docs = []
    for c in args.corpus:
        docs.extend(load_corpus(Path(c).resolve()))
    if not docs:
        print("没有读到语料", file=sys.stderr)
        return 2

    # ---- 全局四字词统计（跨文档聚合，单文档内高频不算"领导口头禅"）----
    per_doc_4, per_doc_5 = [], []
    for _label, body in docs:
        per_doc_4.append(ngram_4char(body))
        per_doc_5.append(ngram_k(body, 5))
    global_4 = Counter()
    global_5 = Counter()
    for c in per_doc_4:
        global_4.update(c)
    for c in per_doc_5:
        global_5.update(c)

    # 至少出现在 2 份不同文档中，降低个案噪声
    doc_presence = Counter()
    for c in per_doc_4:
        for w in c:
            doc_presence[w] += 1

    # 碎片分离：去掉长专名/术语的切片（千家万户水 / 质净化厂 / 户水管家 …）
    kept, frags = split_fragments(global_4, global_5, ratio=args.frag_ratio)

    def _row(w, n):
        return {"word": w, "count": n, "docs": doc_presence[w],
                "is_idiom": w in idiom_seed}

    pool = [_row(w, n) for w, n in kept if n >= args.min_4char_freq and doc_presence[w] >= 2]
    pool.sort(key=lambda d: (-d["count"], d["word"]))
    idioms = [d for d in pool if d["is_idiom"]][: args.top_4char]
    domain = [d for d in pool if not d["is_idiom"]][: args.top_4char]

    idiom_set = {d["word"] for d in idioms}
    four_char = idioms  # 保持旧字段名兼容下游

    if args.only_4char:
        four_out = {"idioms": idioms, "domain_terms": domain,
                    "fragments_dropped": sorted(frags, key=lambda d: -d["count"])[:80]}
        Path(args.out + ".json").write_text(
            json.dumps(four_out, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"四字词：成语 {len(idioms)} 条 / 业务术语 {len(domain)} 条 / "
              f"丢弃碎片 {len(frags)} 条 → {args.out}.json")
        return 0

    # ---- 金句候选 ----
    seen = set()
    cands = []
    for label, body in docs:
        for s in split_sentences(body):
            key = s.rstrip("。！？!?")
            if len(key) < 8 or len(key) > 80:
                continue
            if key in seen:
                continue
            seen.add(key)
            hits = sum(1 for i in range(len(key) - 3) if key[i:i + 4] in idiom_set)
            sc, tags = score_sentence(s, hits)
            if sc < args.min_score:
                continue
            cands.append({"text": key, "score": sc, "tags": tags, "source": label})

    cands.sort(key=lambda d: (-d["score"], -len(d["tags"]), d["text"]))
    cands = cands[: args.top]

    # ---- 落盘 ----
    payload = {"corpus_files": args.corpus, "doc_count": len(docs),
               "idioms": idioms, "domain_terms": domain,
               "fragments_dropped": sorted(frags, key=lambda d: -d["count"])[:80],
               "candidates": cands}
    Path(args.out + ".json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    by_tag = defaultdict(list)
    for c in cands:
        for t in c["tags"]:
            by_tag[t].append(c)

    lines = [f"# 金句候选（{len(cands)} 条，来自 {len(docs)} 份样本）", "",
             "> ⚠️ 脚本输出仅为**候选**，须人工复核后写入 voice_<领导>.md。", "",
             "## 全量候选（按分数降序，去重）", ""]
    for i, c in enumerate(cands[:120], 1):
        lines.append(f"{i:>3}. [{c['score']}] {c['text']}")
        lines.append(f"     标签：{'/'.join(c['tags'])}　出处：{c['source'].split('/')[-1]}")
    lines.append("")

    for t in ("格言反差", "号召动员", "比喻意象", "对仗结构", "排比结构", "递进转折", "引语包装"):
        items = by_tag.get(t, [])
        if not items:
            continue
        lines.append(f"## {t}（{len(items)}）")
        for c in items[:40]:
            lines.append(f"- [{c['score']}] {c['text']}   `{c['source'].split('/')[-1]}`")
        lines.append("")
    lines.append(f"## 高频四字词 / 成语（{len(idioms)}）")
    lines.append("、".join(f"{d['word']}({d['count']})" for d in idioms))
    lines.append("")
    lines.append(f"## 高频业务术语（{len(domain)}，非成语，供战略语汇参考）")
    lines.append("、".join(f"{d['word']}({d['count']})" for d in domain))
    lines.append("")
    lines.append(f"## 已剔除的长词碎片（{len(frags)}，前 40）")
    lines.append("、".join(f"{d['word']}({d['count']}→{d['ext_count']})"
                          for d in sorted(frags, key=lambda x: -x["count"])[:40]))
    Path(args.out + ".md").write_text("\n".join(lines), encoding="utf-8")

    print(f"样本 {len(docs)} 份；金句候选 {len(cands)} 条；"
          f"成语 {len(idioms)} 条；业务术语 {len(domain)} 条；剔除碎片 {len(frags)} 条")
    print(f"→ {args.out}.json / {args.out}.md")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
