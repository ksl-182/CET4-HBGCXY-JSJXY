"""CET-4 四级 1500 词作文生成器
从 GitHub 获取四级核心词汇，生成 10 篇真实作文。
用 pandoc 输出中文版和英文版两份 Word 文档。
用法: python generate_cet4_essays.py
"""
import csv
import json
import os
import random
import subprocess
import sys
import io
from dataclasses import dataclass, field
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

PANDOC = os.path.expanduser("~/pandoc/pandoc.exe")

# ── 颜色方案 ────────────────────────────────────────────────────────
COLORS = {
    "keyword": "color:#DC143C",     # 深红 - 核心词汇
    "pattern": "color:#0066CC",     # 蓝色 - 句型
    "transition": "color:#009900",  # 绿色 - 过渡词
    "phrase": "color:#FF8C00",      # 橙色 - 高频短语
}

CATEGORY_CN = {
    "keyword": "核心词汇",
    "pattern": "关键句型",
    "transition": "过渡词",
    "phrase": "高频短语",
}


@dataclass
class Word:
    word: str
    meaning: str


# ── 1. 词汇获取 ────────────────────────────────────────────────────

def load_vocab() -> list[Word]:
    """从本地缓存读取词汇"""
    print("[1/3] 获取四级词汇...")
    tmp = os.environ.get("TEMP", "/tmp")

    # CSV 高频词
    csv_words = []
    csv_path = os.path.join(tmp, "cet4_high_freq.csv")
    with open(csv_path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            w = row.get("word", "").strip().lower()
            if w and w.isalpha() and len(w) > 1:
                csv_words.append(Word(word=w, meaning=""))
    print(f"  CSV: {len(csv_words)}")

    # JSON 完整词表（补充释义）
    json_map = {}
    json_dir = os.path.join(tmp, "cet4_json")
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        path = os.path.join(json_dir, f"{letter}.json")
        try:
            with open(path, encoding="utf-8") as f:
                for entry in json.load(f):
                    w = entry.get("word", "").strip().lower()
                    m = entry.get("mean", "").strip()
                    if w and m:
                        json_map[w] = m
        except Exception:
            pass
    print(f"  JSON: {len(json_map)}")

    # 合并
    seen = set()
    result = []
    for w in csv_words:
        key = w.word.lower()
        if key not in seen:
            seen.add(key)
            result.append(Word(word=w.word, meaning=json_map.get(key, "")))
    # 补充
    for w, m in json_map.items():
        if w not in seen and len(result) < 1500:
            seen.add(w)
            result.append(Word(word=w, meaning=m))
    result = result[:1500]
    print(f"  合并: {len(result)}")
    return result


# ── 2. 词汇分配 ────────────────────────────────────────────────────

TOPICS = [
    ("教育", "Education"),
    ("科技", "Technology"),
    ("环保", "Environment"),
    ("健康", "Health"),
    ("文化", "Culture"),
    ("社交媒体", "Social Media"),
    ("就业", "Employment"),
    ("志愿服务", "Volunteering"),
    ("传统", "Traditions"),
    ("个人成长", "Personal Growth"),
]


def distribute_words(words: list[Word], seed: int = 42) -> dict[str, list[Word]]:
    shuffled = list(words)
    random.seed(seed)
    random.shuffle(shuffled)
    n = len(shuffled)
    per = n // 10
    rem = n % 10
    result = {}
    idx = 0
    for i, (_, en) in enumerate(TOPICS):
        cnt = per + (1 if i < rem else 0)
        result[en] = shuffled[idx:idx + cnt]
        idx += cnt
    return result


# ── 3. 作文生成（真实作文，非单词列表）──────────────────────────────

def s(text: str, cat: str = "normal", meaning: str = "") -> tuple:
    """创建一个文本片段"""
    return (text, cat, meaning)


def w(word: Word, cat: str = "keyword") -> tuple:
    """创建一个词汇片段"""
    return (word.word, cat, word.meaning)


def build_essay_english(topic: str, words: list[Word]) -> list[list[tuple]]:
    """生成英文作文段落：真实句子 + 嵌入词汇"""
    random.shuffle(words)
    n = len(words)
    # 分组：开头、主体、结尾
    g1 = words[: n // 3]
    g2 = words[n // 3: 2 * n // 3]
    g3 = words[2 * n // 3:]

    paras = []

    # ── 开头段 ──
    p1 = []
    p1.append(s("It is widely acknowledged that ", "pattern", "人们普遍认为"))
    p1.append(s(topic.lower()))
    p1.append(s(" plays a crucial role in ", "phrase", "在……中起关键作用"))
    p1.append(s("modern society. "))
    p1.append(s("Furthermore, ", "transition", "此外"))
    p1.append(s("this issue has attracted widespread attention from "))
    # 嵌入 5-8 个词到一个句子
    chunk1 = g1[:8]
    for i, word in enumerate(chunk1):
        p1.append(w(word))
        if i < len(chunk1) - 2:
            p1.append(s(", "))
        elif i == len(chunk1) - 2:
            p1.append(s(" and "))
    p1.append(s(", which "))
    p1.append(s("has a profound impact on ", "phrase", "对……有深远影响"))
    p1.append(s("our daily life. "))

    # 第二句嵌入更多词
    chunk2 = g1[8:16] if len(g1) > 8 else []
    if chunk2:
        p1.append(s("In particular, ", "transition", "特别是"))
        p1.append(s("many people "))
        p1.append(s("pay attention to ", "phrase", "注意"))
        p1.append(s("the relationship between "))
        for i, word in enumerate(chunk2):
            p1.append(w(word))
            if i < len(chunk2) - 2:
                p1.append(s(", "))
            elif i == len(chunk2) - 2:
                p1.append(s(" and "))
        p1.append(s(", because "))
        p1.append(s("it is essential for us to ", "pattern", "对我们来说……是必不可少的"))
        p1.append(s("understand its significance."))

    paras.append(p1)

    # ── 主体段 ──
    p2 = []
    p2.append(s("There is no denying that ", "pattern", "不可否认的是"))
    chunk3 = g2[:6]
    for i, word in enumerate(chunk3):
        p2.append(w(word))
        if i < len(chunk3) - 2:
            p2.append(s(", "))
        elif i == len(chunk3) - 2:
            p2.append(s(" and "))
    p2.append(s(" are closely related to each other. "))
    p2.append(s("For example, ", "transition", "例如"))
    p2.append(s("we should "))
    p2.append(s("attach great importance to ", "phrase", "高度重视"))
    chunk4 = g2[6:12] if len(g2) > 6 else []
    if chunk4:
        for i, word in enumerate(chunk4):
            p2.append(w(word))
            if i < len(chunk4) - 2:
                p2.append(s(", "))
            elif i == len(chunk4) - 2:
                p2.append(s(" and "))
    p2.append(s(". "))
    p2.append(s("Moreover, ", "transition", "而且"))
    p2.append(s("the reason why "))
    chunk5 = g2[12:18] if len(g2) > 12 else g2[6:]
    if chunk5:
        for i, word in enumerate(chunk5[:4]):
            p2.append(w(word))
            if i < min(len(chunk5), 4) - 2:
                p2.append(s(", "))
            elif i == min(len(chunk5), 4) - 2:
                p2.append(s(" and "))
    p2.append(s(" matter so much "))
    p2.append(s("is that ", "pattern", "……的原因是"))
    p2.append(s("they "))
    p2.append(s("lay a solid foundation for ", "phrase", "为……打下坚实基础"))
    p2.append(s("future development. "))

    # 第二主体句
    chunk6 = g2[18:] if len(g2) > 18 else []
    if chunk6:
        p2.append(s("From my perspective, ", "pattern", "在我看来"))
        p2.append(s("we should "))
        p2.append(s("spare no effort to ", "phrase", "不遗余力地"))
        p2.append(s("promote "))
        for i, word in enumerate(chunk6[:5]):
            p2.append(w(word))
            if i < min(len(chunk6), 5) - 2:
                p2.append(s(", "))
            elif i == min(len(chunk6), 5) - 2:
                p2.append(s(" and "))
        p2.append(s(", because "))
        p2.append(s("nothing is more important than ", "pattern", "没有什么比……更重要"))
        p2.append(s("taking action."))

    paras.append(p2)

    # ── 结尾段 ──
    p3 = []
    p3.append(s("In conclusion, ", "transition", "总之"))
    chunk7 = g3[:8]
    for i, word in enumerate(chunk7):
        p3.append(w(word))
        if i < len(chunk7) - 2:
            p3.append(s(", "))
        elif i == len(chunk7) - 2:
            p3.append(s(" and "))
    p3.append(s(" are all important factors. "))
    p3.append(s("Only by doing this can we ", "pattern", "只有通过……我们才能"))
    p3.append(s("make a real difference ", "phrase", "产生影响"))
    p3.append(s("and "))
    p3.append(s("adapt to the changing world. ", "phrase", "适应变化的世界"))
    p3.append(s("Therefore, ", "transition", "因此"))
    p3.append(s("let us "))
    p3.append(s("be aware of ", "phrase", "意识到"))
    p3.append(s("the challenges ahead and work together for a better future."))

    paras.append(p3)
    return paras


def build_essay_chinese(topic: str, words: list[Word]) -> list[list[tuple]]:
    """生成中文作文段落：真实句子 + 嵌入词汇"""
    random.shuffle(words)
    n = len(words)
    g1 = words[: n // 3]
    g2 = words[n // 3: 2 * n // 3]
    g3 = words[2 * n // 3:]

    paras = []

    # ── 开头段 ──
    p1 = []
    p1.append(s("人们普遍认为，", "pattern", "It is widely acknowledged that"))
    p1.append(s(f"{topic}在现代社会中"))
    p1.append(s("起着关键作用。", "phrase", "play a crucial role in"))
    p1.append(s("此外，", "transition", "Furthermore"))
    p1.append(s("这一问题已经引起了广泛关注，"))
    chunk1 = g1[:8]
    for i, word in enumerate(chunk1):
        p1.append(s(word.meaning if word.meaning else word.word, "keyword", word.word))
        if i < len(chunk1) - 2:
            p1.append(s("、"))
        elif i == len(chunk1) - 2:
            p1.append(s("和"))
    p1.append(s("等都是我们需要关注的方面。"))
    p1.append(s("这些问题", "phrase", "have a profound impact on"))
    p1.append(s("对我们的日常生活有深远影响。"))

    chunk2 = g1[8:16] if len(g1) > 8 else []
    if chunk2:
        p1.append(s("特别是，", "transition", "In particular"))
        p1.append(s("许多人"))
        p1.append(s("关注", "phrase", "pay attention to"))
        p1.append(s("它们之间的关系，因为"))
        p1.append(s("对我们来说理解其意义是必不可少的。", "pattern", "It is essential for us to"))
        for i, word in enumerate(chunk2[:6]):
            p1.append(s(word.meaning if word.meaning else word.word, "keyword", word.word))
            if i < min(len(chunk2), 6) - 2:
                p1.append(s("、"))
            elif i == min(len(chunk2), 6) - 2:
                p1.append(s("等。"))

    paras.append(p1)

    # ── 主体段 ──
    p2 = []
    p2.append(s("不可否认的是，", "pattern", "There is no denying that"))
    chunk3 = g2[:6]
    for i, word in enumerate(chunk3):
        p2.append(s(word.meaning if word.meaning else word.word, "keyword", word.word))
        if i < len(chunk3) - 2:
            p2.append(s("、"))
        elif i == len(chunk3) - 2:
            p2.append(s("和"))
    p2.append(s("密切相关。"))
    p2.append(s("例如，", "transition", "For example"))
    p2.append(s("我们应该"))
    p2.append(s("高度重视", "phrase", "attach great importance to"))
    chunk4 = g2[6:12] if len(g2) > 6 else []
    if chunk4:
        for i, word in enumerate(chunk4):
            p2.append(s(word.meaning if word.meaning else word.word, "keyword", word.word))
            if i < len(chunk4) - 2:
                p2.append(s("、"))
            elif i == len(chunk4) - 2:
                p2.append(s("和"))
        p2.append(s("。"))
    p2.append(s("而且，", "transition", "Moreover"))
    p2.append(s("其原因在于，", "pattern", "The reason why ... is that"))
    p2.append(s("它们"))
    p2.append(s("为未来发展打下坚实基础。", "phrase", "lay a solid foundation for"))

    chunk5 = g2[12:18] if len(g2) > 12 else g2[6:]
    if chunk5:
        p2.append(s("在我看来，", "pattern", "From my perspective"))
        p2.append(s("我们应该"))
        p2.append(s("不遗余力地", "phrase", "spare no effort to"))
        p2.append(s("推动"))
        for i, word in enumerate(chunk5[:5]):
            p2.append(s(word.meaning if word.meaning else word.word, "keyword", word.word))
            if i < min(len(chunk5), 5) - 2:
                p2.append(s("、"))
            elif i == min(len(chunk5), 5) - 2:
                p2.append(s("和"))
        p2.append(s("的发展，因为"))
        p2.append(s("没有什么比行动更重要。", "pattern", "Nothing is more important than"))

    paras.append(p2)

    # ── 结尾段 ──
    p3 = []
    p3.append(s("总之，", "transition", "In conclusion"))
    chunk6 = g3[:8]
    for i, word in enumerate(chunk6):
        p3.append(s(word.meaning if word.meaning else word.word, "keyword", word.word))
        if i < len(chunk6) - 2:
            p3.append(s("、"))
        elif i == len(chunk6) - 2:
            p3.append(s("和"))
    p3.append(s("都是重要因素。"))
    p3.append(s("只有这样，我们才能", "pattern", "Only by doing this can we"))
    p3.append(s("产生真正的影响", "phrase", "make a real difference"))
    p3.append(s("并"))
    p3.append(s("适应变化的世界。", "phrase", "adapt to the changing world"))
    p3.append(s("因此，", "transition", "Therefore"))
    p3.append(s("让我们"))
    p3.append(s("意识到", "phrase", "be aware of"))
    p3.append(s("前方的挑战，共同努力创造更美好的未来。"))

    paras.append(p3)
    return paras


# ── 4. Markdown 生成 ───────────────────────────────────────────────

def span(text: str, cat: str) -> str:
    """生成带颜色的 HTML span"""
    if cat == "normal":
        return text
    style = COLORS.get(cat, COLORS["keyword"])
    return f'<span style="{style}">**{text}**</span>'


def segments_to_md(segments: list[tuple], is_english: bool) -> str:
    """将片段列表转为 Markdown 行"""
    parts = []
    for text, cat, meaning in segments:
        if cat == "normal":
            parts.append(text)
        elif is_english:
            # 英文文档：只显示英文，不加中文注释
            parts.append(span(text, cat))
        else:
            # 中文文档：释义(英文词)
            display = f"{text}({meaning})" if meaning else text
            parts.append(span(display, cat))
    return "".join(parts)


def build_markdown(essays: list, is_english: bool) -> str:
    """生成完整 Markdown 文档"""
    lang = "英文" if is_english else "中文"
    lines = []
    lines.append(f"# CET-4 四级核心词汇作文集（{lang}版）\n")
    lines.append(f"生成日期：{datetime.now().strftime('%Y-%m-%d')}\n")
    lines.append("包含 10 篇作文，覆盖 1500+ 四级核心词汇\n")

    # 颜色图例
    lines.append("## 颜色图例\n")
    for cat, style in COLORS.items():
        label = CATEGORY_CN[cat]
        lines.append(f'- <span style="{style}">**■ {label}**</span>')
    lines.append("")

    # 10 篇作文
    for i, (topic_cn, topic_en, paras_cn, paras_en, words) in enumerate(essays, 1):
        topic = topic_en if is_english else topic_cn
        lines.append(f"## Essay {i}: {topic}\n")

        paras = paras_en if is_english else paras_cn
        for para in paras:
            lines.append(segments_to_md(para, is_english))
            lines.append("")

        # 词汇索引表
        lines.append(f"### {'Word Index' if is_english else '词汇索引'}\n")
        lines.append("| Word | Meaning | Category |")
        lines.append("|------|---------|----------|")
        for word in words:
            lines.append(f"| {word.word} | {word.meaning if word.meaning else '-'} | 核心词汇 |")
        lines.append("")

    return "\n".join(lines)


# ── 5. 主入口 ──────────────────────────────────────────────────────

def main():
    words = load_vocab()

    print("[2/3] 分配词汇到 10 个话题...")
    topic_words = distribute_words(words)

    print("[3/3] 生成作文...")
    essays = []
    for cn, en in TOPICS:
        wds = topic_words[en]
        paras_cn = build_essay_chinese(cn, wds)
        paras_en = build_essay_english(en, wds)
        essays.append((cn, en, paras_cn, paras_en, wds))
        print(f"  [OK] {cn} ({en}) - {len(wds)} words")

    # 生成 Markdown
    print("\n生成文档...")
    md_cn = build_markdown(essays, is_english=False)
    md_en = build_markdown(essays, is_english=True)

    # 写入临时 Markdown
    tmp = os.environ.get("TEMP", "/tmp")
    md_cn_path = os.path.join(tmp, "cet4_cn.md")
    md_en_path = os.path.join(tmp, "cet4_en.md")
    with open(md_cn_path, "w", encoding="utf-8") as f:
        f.write(md_cn)
    with open(md_en_path, "w", encoding="utf-8") as f:
        f.write(md_en)

    # pandoc 转 docx
    home = os.path.expanduser("~")
    out_cn = os.path.join(home, "CET4_Chinese.docx")
    out_en = os.path.join(home, "CET4_English_v2.docx")

    subprocess.run([PANDOC, md_cn_path, "-o", out_cn], check=True)
    print(f"  [OK] {out_cn}")

    subprocess.run([PANDOC, md_en_path, "-o", out_en], check=True)
    print(f"  [OK] {out_en}")

    print(f"\nDone! {len(words)} words, 10 essays, 2 documents.")


if __name__ == "__main__":
    main()
