# -*- coding: utf-8 -*-
"""
灯具设计师理论题库 PDF 解析器 v2
输入: tools/_raw_all.txt (pypdfium2 抽取, 带 ===== PAGE n ===== 标记)
输出: tools/_parsed_raw.json + tools/_parse_report.txt

策略:
  1) 顺序扫描期望题号 n=1..800, 在任意行内定位 "n." 起点(支持一行多题/题号无点)
  2) 块内按 "A./B./C./D." 切选项; 找不到字母标号时把题干后的连续短行当作选项
  3) 从题干括号抽答案, 并把答案括弧替换为占位括弧
  4) 答案长度 >=2 -> 多选; 对/错 -> 判断; 否则单选
"""
import re
import json
from pathlib import Path
from collections import Counter

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
RAW = HERE / "_raw_all.txt"
OUT = HERE / "_parsed_raw.json"
REPORT = HERE / "_parse_report.txt"

QNUM_TYPE = re.compile(r"^\s*(单选题|多选题|判断题)\s*")
OPT_START = re.compile(r"^\s*([A-D])\s*[.．、]\s*(.*)$")
ANS_RE = re.compile(r"[（(]\s*([A-Da-d]{1,4}|对|错|正确|错误|√|×|T|F)\s*[.．。、]?\s*[）)]")
JUDGE_ANS = {"对": "T", "正确": "T", "√": "T", "T": "T", "t": "T",
             "错": "F", "错误": "F", "×": "F", "F": "F", "f": "F"}
TOTAL = 800

report = []


def log(*a):
    report.append(" ".join(str(x) for x in a))


def clean(s: str) -> str:
    s = s.replace("\u00a0", " ").replace("\u3000", " ")
    s = re.sub(r"[ \t]+", " ", s).strip()
    prev = None
    while prev != s:
        prev = s
        s = re.sub(r"(?<=[\u3400-\u9fff\u3000-\u303f\uff00-\uffef])[ ]+", "", s)
        s = re.sub(r"[ ]+(?=[\u3400-\u9fff\u3000-\u303f\uff00-\uffef])", "", s)
    s = s.strip()
    # PDF 两端对齐会在 "℃" 后留下空格，删除中文空格规则会把两段词粘死
    # （如 "金属部件 65℃ 非金属部件 75℃" -> "金属部件65℃非金属部件75℃"），这里补回分隔符
    s = re.sub(r"℃(?=[\u4e00-\u9fff])", "℃，", s)
    return s


def split_tail_answer(text, ans):
    """把粘在选项末尾的答案字母剥出来。

    题库 PDF 里不少题把答案字母排到了最后一个选项的尾部，例如
    "D.5−10 B"、"C.分层实体制造成型LOM C"、"D.以上答案都不对B"。
    只在「前面是空白或中文字符」且「字母集合等于答案集合」时才剥离，
    避免误伤 "AutoCAD"、"Copy CAD"、"1921,ACSIC" 这类正常选项文本。
    返回 (新文本, 剥离出的字母串或 None)。
    """
    m = re.search(r"([A-D]{1,4})$", text)
    if not m:
        return text, None
    letters = m.group(1)
    head = text[:m.start()]
    if not head:
        return text, None
    prev = head[-1]
    if not (prev in " \u3000。，、；：" or "\u4e00" <= prev <= "\u9fff"):
        return text, None
    if not ans:
        return head.rstrip(" \u3000"), letters
    if set(letters) != set(ans):
        return text, None
    return head.rstrip(" \u3000"), letters


def load_lines():
    page = 0
    out = []
    for ln in RAW.read_text(encoding="utf-8").split("\n"):
        m = re.match(r"^=====\s*PAGE\s+(\d+)\s*=====$", ln.strip())
        if m:
            page = int(m.group(1))
            continue
        if ln.strip():
            out.append((ln.rstrip(), page))
    return out


def qstart_patterns(n):
    nb = r"(?<![0-9])" + str(n) + r"(?![0-9])"
    return [
        re.compile(nb + r"\s*[.．、]\s*"),
        re.compile(nb + r"(?=\s*[“\"‘'（(【])"),
        re.compile(nb + r"\s+(?=(?:单选题|多选题|判断题))"),
    ]


def find_qstart(s, n):
    best = None
    for p in qstart_patterns(n):
        m = p.search(s)
        if m and (best is None or m.start() < best.start()):
            best = m
    return best


def split_blocks(lines):
    """顺序扫描，支持一行多题。

    关键：匹配到题号后只把「题号之前」的文本收进当前块，剩余部分留在 rest 里
    交给下一轮继续找下一个题号——绝不能先把整个 rest 收进当前块，否则当同一行
    后面还跟着下一题时（如 "...支撑。 625.（错）STL"），当前块就会把下一题的开头
    也吞进去，造成题干被污染 + 文本重复。
    """
    blocks = []
    cur = None
    expected = 1
    for ln, page in lines:
        rest = ln
        guard = 0
        while True:
            guard += 1
            if guard > 40:
                break
            m = find_qstart(rest, expected) if expected <= TOTAL else None
            if m is None:
                if cur is not None and rest.strip():
                    cur[2].append(rest)
                break
            pre = rest[:m.start()]
            if cur is not None and pre.strip():
                cur[2].append(pre)
            if cur is not None:
                blocks.append(cur)
            cur = [expected, page, []]
            expected += 1
            rest = rest[m.end():]
            if not rest.strip():
                break
    if cur is not None:
        blocks.append(cur)
    return blocks


def split_options(block_lines):
    stem, opts, cur = [], [], None
    for ln in block_lines:
        m = OPT_START.match(ln)
        if m:
            want = "ABCD"[len(opts)] if len(opts) < 4 else None
            if m.group(1) == want:
                cur = [m.group(1), [m.group(2)]]
                opts.append(cur)
                continue
        if cur is not None:
            cur[1].append(ln)
        else:
            stem.append(ln)
    return stem, opts


def join_lines(parts):
    return "".join(clean(p) for p in parts if p is not None).strip()


def fix_text(s):
    s = re.sub(r"(?<![0-9A-Za-z])D\s*打印", "3D打印", s)
    s = re.sub(r"(?<![A-Za-z])FD\s*工艺", "FDM工艺", s)
    return s


def parse_block(num, page, raw_lines):
    rec = {"no": num, "page": page, "type": None, "stem": "", "options": [],
           "answer": [], "issues": [], "prefix": None}
    lines = list(raw_lines)
    if lines:
        # PDF 里有 "418..单选题 ..." 这种双句点，先吃掉多余的标点再匹配题型前缀
        lines[0] = re.sub(r"^[.．、]\s*", "", lines[0])
        m = QNUM_TYPE.match(lines[0])
        if m:
            rec["prefix"] = {"单选题": "single", "多选题": "multi", "判断题": "judge"}[m.group(1)]
            lines[0] = lines[0][m.end():]

    stem_lines, opts = split_options(lines)
    stem_raw = join_lines(stem_lines)

    # 判断题: (对)/(错) 开头
    judge_hit = None
    m = re.match(r"^\s*[（(]\s*(对|错|正确|错误|√|×|T|F|t|f)\s*[）)]\s*", stem_raw)
    if m:
        judge_hit = JUDGE_ANS[m.group(1).strip()]
        stem_raw = stem_raw[m.end():]

    ans = None
    if judge_hit:
        ans = [judge_hit]
        rec["type"] = "judge"
    else:
        m = ANS_RE.search(stem_raw)
        if m:
            v = m.group(1).strip()
            if re.fullmatch(r"[A-Da-d]{1,4}", v):
                ans = sorted(set(v.upper()))
            else:
                ans = [JUDGE_ANS[v]]
                rec["type"] = "judge"
            stem_raw = stem_raw[:m.start()] + "（　　）" + stem_raw[m.end():]

    if rec["type"] != "judge":
        if ans and len(ans) >= 2:
            rec["type"] = "multi"
        elif ans:
            rec["type"] = "single"

    # 没有字母标号的选项: 把题干之后的短行当选项
    if not opts and rec["type"] in ("single", "multi"):
        tail = [x for x in stem_lines[1:] if x.strip()]
        if len(tail) >= 3:
            head = join_lines(stem_lines[:1])
            if ANS_RE.search(head):
                stem_raw = head
                ans_m = ANS_RE.search(stem_raw)
                v = ans_m.group(1).strip()
                if re.fullmatch(r"[A-Da-d]{1,4}", v):
                    ans = sorted(set(v.upper()))
                    rec["type"] = "multi" if len(ans) >= 2 else "single"
                    stem_raw = stem_raw[:ans_m.start()] + "（　　）" + stem_raw[ans_m.end():]
                opts = [[k, [t]] for k, t in zip("ABCD", tail)]

    options = [{"key": k, "text": join_lines(t)} for k, t in opts]

    # 答案字母粘在末选项末尾: 没答案时回填, 有答案时剥离重复字母
    if options and rec["type"] != "judge":
        text, letters = split_tail_answer(options[-1]["text"], ans)
        if letters and text != options[-1]["text"]:
            options[-1]["text"] = text
            if not ans:
                ans = sorted(set(letters))
                rec["type"] = "multi" if len(ans) >= 2 else "single"
                rec["issues"].append(f"答案从末选项末尾回填({letters})")
            else:
                rec["issues"].append(f"从末选项末尾剥离重复答案字母({letters})")

    if rec["type"] == "judge":
        options = [{"key": "T", "text": "正确"}, {"key": "F", "text": "错误"}]

    stem_raw = fix_text(stem_raw)
    options = [{"key": o["key"], "text": fix_text(o["text"])} for o in options]

    # ---- 校验 ----
    iss = rec["issues"]
    if rec["type"] is None:
        iss.append("无法判定题型")
    if rec["type"] in ("single", "multi"):
        if len(options) != 4:
            iss.append(f"选项数={len(options)}")
        if not ans:
            iss.append("未找到答案")
        elif not set(ans) <= {o["key"] for o in options}:
            iss.append(f"答案 {''.join(ans)} 不在选项 {''.join(o['key'] for o in options)} 中")
        if rec["prefix"] == "multi" and len(ans or []) == 1:
            iss.append("前缀=多选但答案仅 1 项")
        if rec["prefix"] == "single" and len(ans or []) > 1:
            iss.append("前缀=单选但答案多项")
    if rec["prefix"] == "judge" and rec["type"] != "judge":
        iss.append("前缀=判断但无对错答案")
    if not stem_raw:
        iss.append("题干为空")

    rec.update({"stem": stem_raw, "options": options, "answer": ans or []})
    return rec


def main():
    lines = load_lines()
    log(f"有效行: {len(lines)}")
    blocks = split_blocks(lines)
    log(f"题块: {len(blocks)} / 期望 {TOTAL}")
    nums = [b[0] for b in blocks]
    log(f"缺失题号: {[n for n in range(1, TOTAL + 1) if n not in set(nums)]}")
    log(f"重复题号: {[n for n, c in Counter(nums).items() if c > 1]}")

    recs = []
    for num, page, rl in blocks:
        recs.append(parse_block(num, page, rl))

    bad = [r for r in recs if r["issues"]]
    log(f"\n有问题: {len(bad)} 题")
    for r in bad:
        log(f"  Q{r['no']:>3} p{r['page']:>3} type={r['type']} ans={r['answer']} opts={len(r['options'])} :: {', '.join(r['issues'])}")
        log(f"      stem={r['stem'][:110]}")
        for o in r["options"]:
            log(f"        {o['key']}. {o['text'][:90]}")

    log("\n题型: " + str(Counter(r["type"] for r in recs)))
    log("答案长度: " + str(Counter(len(r["answer"]) for r in recs)))

    OUT.write_text(json.dumps(recs, ensure_ascii=False, indent=1), encoding="utf-8")
    REPORT.write_text("\n".join(report), encoding="utf-8")
    print("\n".join(report))
    print(f"\n写出 {OUT}")


if __name__ == "__main__":
    main()
