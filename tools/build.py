# -*- coding: utf-8 -*-
"""
构建脚本：把 tools/_parsed_raw.json 加工成站点数据 + 人工校对表。

步骤
  1) 应用人工补丁（PATCHES）
  2) 严格去重：规范化题干 + 选项内容集合 + 映射到选项内容的答案 三者全同
  3) 章节归类：关键词打分 + 未分类题按邻近题块补齐
  4) 输出 data/questions.json（SPEC 契约）与 docs/题库校对表.xlsx

用法: python tools/build.py
"""
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
RAW = HERE / "_parsed_raw.json"
OUT_JSON = ROOT / "data" / "questions.json"
OUT_XLSX = ROOT / "docs" / "题库校对表.xlsx"

TITLE = "第四届全国工业设计职业技能大赛 灯具设计师赛项理论题库"
VERSION = "2024-shandong-r1"


# ---------------------------------------------------------------- 人工补丁
# 只放「算法无法可靠修好、且已确认过原文」的题。每处都写清依据，便于复核。
PATCHES = {
    # --- 源 PDF 答案括号为空，按标准教材答案补齐，待用户确认 ---
    89: {"answer": "A", "flag": "ai-answer", "note": "原PDF答案空缺；3D打印在建筑中最广泛用途=建筑装饰品和建筑模型生产，补A"},
    211: {"answer": "A", "flag": "ai-answer", "note": "原PDF答案空缺；风格派代表人物=杜斯伯格(Theo van Doesburg)，补A"},
    259: {"answer": "B", "flag": "ai-answer", "note": "原PDF答案空缺；演绎推理常用方式含三段论/线性/条件推理，判断推理不属于，补B"},
    345: {"answer": "D", "flag": "ai-answer", "note": "原PDF答案空缺；截平面与圆锥轴线斜交且平行一条素线→抛物线，补D"},
    347: {"answer": "B", "flag": "ai-answer", "note": "原PDF答案空缺；球面直径/半径标注在∅、R前加S，补B"},
    348: {"answer": "B", "flag": "ai-answer", "note": "原PDF答案空缺；两立体相交表面交线=相贯线，补B"},
    352: {"answer": "A", "flag": "ai-answer", "note": "原PDF答案空缺；正等轴测图简化轴向伸缩系数取1（理论值0.82不在选项中），补A"},
    371: {"answer": "D", "flag": "ai-answer", "note": "原PDF答案空缺；倾斜结构实形用斜视图表达，补D"},
    412: {"answer": "D", "flag": "ai-answer", "note": "原PDF答案空缺；GB/T120.2为内螺纹圆柱销，补D"},

    # --- 选项结构在 PDF 里被破坏，按原文重建 ---
    265: {
        "options": ["功能主义", "形式主义", "孟非斯", "工艺美术运动"],
        "answer": "C",
        "note": "原PDF第4个选项字母误排为C（重复），按其文本应为D；答案C(孟非斯)",
    },
    116: {
        "options": [
            "可快速构建瓶状或中空零件",
            "制造系统不用于办公环境，因为会产生有毒化学物质",
            "原材料以卷轴丝的形式提供，易于搬运和快速更换",
        ],
        "answer": "B",
        "flag": "review",
        "note": "原PDF选项A/B被换行拆断、C吞并了两条选项，原文只剩3条（缺D项）；已按语义还原",
    },
    538: {
        "options": [
            "Pro/Engineer（简称Pro/E，现改名为Creo）是美国参数技术公司（Parametric Technology Corporation，简称PTC）的高端产品。PTC公司提出的参数化、基于特征、全相关的概念改变了机械CAD/CAE/CAM的传统观念，这种概念已成为标准",
            "Pro/Engineer软件能将设计至生产全过程集成到一起，让所有的用户能够同时进行同一产品的设计制造工作，即实现所谓的串行工程",
            "Cimatron系统是以色列Cimatron公司的CAD/CAE/CAM产品。该系统提供了比较灵活的用户界面，优良的三维造型、工程绘图，全面的数控加工，各种通用、专用数据接口以及集成化的产品数据管理",
            "Cimatron系统在国际上的模具制造业备受欢迎",
        ],
        "answer": "ACD",
        "flag": "review",
        "note": "原PDF此题为双栏排版，文字层交错错乱；已按乱序文本还原4个选项，答案为ACD（B错在串行工程应为并行工程）",
    },
    46: {
        "flag": "broken",
        "note": "题目要求识别灯具控制装置图形符号，但PDF文字层只有符号字体乱码（8盾/8v/8低O/不像八），无法还原图形，建议删除此题",
    },
    113: {
        "flag": "review",
        "note": "原PDF只有3个选项且A写作“AB塑料丝”（疑为ABS漏S）；未改动原文",
    },
    115: {
        "flag": "review",
        "note": "原PDF只有3个选项（SLS / 3DP / FDM），未改动原文",
    },
    60: {
        "flag": "broken",
        "note": "原PDF此题题干是乱码（1+D62:I521+D62:I80），实为第1题(STL用一系列平面逼近模型)的重复副本，题干已损坏无法还原；建议删除此题",
    },
    284: {
        "flag": "broken",
        "note": "原PDF此题题干与选项明显不匹配（题干问形态配置手段，选项却是透明工件/大尺寸件/反光工件等），选项属于另一道题；建议删除此题",
    },
    84: {"flag": "review", "note": "原PDF选项C与D文本相同（都为“透射深度大”），排版缺陷，保留原样"},
    224: {"flag": "review", "note": "原PDF选项C与D文本相同，排版缺陷，保留原样"},
    252: {"flag": "review", "note": "原PDF选项B与D文本相同（都为亨利·德雷夫斯），排版缺陷，保留原样"},
    468: {"flag": "review", "note": "原PDF选项A与B文本相同（都为财务管理），排版缺陷，保留原样"},
}

# ---------------------------------------------------------------- 章节定义
CHAPTERS = [
    ("c1", "3D打印与增材制造"),
    ("c2", "设计史与设计理论"),
    ("c3", "色彩与构成"),
    ("c4", "机械制图与公差"),
    ("c5", "灯具与电气安全"),
    ("c6", "CAD·PDM与软件"),
    ("misc", "综合与其他"),
]

KEYWORDS = {
    "c1": [
        ("3D打印", 3), ("增材", 3), ("打印", 2), ("熔融沉积", 3), ("挤压成型", 2), ("光固化", 3),
        ("烧结", 3), ("SLS", 3), ("SLA", 3), ("FDM", 3), ("LOM", 3), ("3DP", 3), ("SLM", 3),
        ("STL", 3), ("喷头", 3), ("支撑", 2), ("分层", 2), ("切片", 3), ("丝材", 3), ("耗材", 2),
        ("三维扫描", 3), ("扫描仪", 2), ("快速成形", 3), ("快速成型", 3), ("成型机", 2),
        ("光敏树脂", 3), ("粉末", 2), ("模型精度", 2), ("打印平台", 3), ("打印头", 3),
    ],
    "c2": [
        ("设计史", 3), ("包豪斯", 3), ("莫里斯", 3), ("工艺美术运动", 3), ("新艺术", 3),
        ("孟菲斯", 3), ("孟非斯", 3), ("风格派", 3), ("构成主义", 3), ("现代主义", 3),
        ("后现代", 3), ("工业设计", 2), ("设计理论", 2), ("审美", 2), ("人机工程", 3),
        ("芝加哥学派", 3), ("维也纳", 2), ("德意志制造联盟", 3), ("罗维", 3), ("拉姆斯", 3),
        ("索特萨斯", 3), ("红蓝椅", 3), ("里特维尔德", 3), ("康定斯基", 3), ("设计流派", 3),
        ("设计师", 2), ("设计风格", 3), ("文化遗产", 2), ("文化自觉", 2), ("费孝通", 3),
        ("装饰", 2), ("美学", 2), ("造型", 1), ("毕加索", 3), ("达芬奇", 3),
        ("柱式", 3), ("哥特式", 3), ("巴洛克", 3), ("起源说", 3), ("新石器", 3),
        ("天工开物", 3), ("考工记", 3), ("营造法式", 3), ("农书", 3),
        ("旧石器", 3), ("形式与功能", 3), ("新产品", 3), ("商业模式", 3), ("功能主义", 2),
    ],
    "c3": [
        ("色彩", 3), ("色相", 3), ("明度", 3), ("饱和度", 3), ("色温", 3), ("三原色", 3),
        ("RGB", 3), ("CMYK", 3), ("分辨率", 3), ("像素", 3), ("色彩模式", 3), ("对比色", 3),
        ("邻近色", 3), ("补色", 3), ("平面设计", 3), ("版式", 3), ("字体", 2), ("标志", 2),
        ("图象", 3), ("图像", 3), ("PHOTOSHOP", 3), ("Illustrator", 3), ("配色", 3),
        ("冷暖", 2), ("光度", 2), ("亮度", 2), ("视觉", 2), ("构图", 2), ("三大构成", 3),
        ("构成", 2), ("平面构成", 3), ("色彩构成", 3), ("立体构成", 3), ("骨格", 3),
        ("骨骼", 3), ("肌理", 3), ("形态", 2), ("形式美", 3), ("渐变", 2), ("发射", 2),
        ("特异", 3), ("重复", 2), ("对比", 1), ("统一性", 2), ("点线面", 3), ("图案", 2),
    ],
    "c4": [
        ("制图", 3), ("视图", 3), ("剖视", 3), ("断面", 3), ("公差", 3), ("配合", 3),
        ("粗糙度", 3), ("螺纹", 3), ("装配", 3), ("尺寸", 2), ("投影", 3), ("相贯", 3),
        ("截交", 3), ("轴测", 3), ("形位公差", 3), ("标准件", 3), ("齿轮", 3), ("键连接", 3),
        ("销", 2), ("弹簧", 3), ("零件图", 3), ("装配图", 3), ("基准", 2), ("表面粗糙度", 3),
        ("H7", 2), ("公差带", 3), ("尺寸线", 3), ("标题栏", 3), ("图样", 3), ("剖切", 3),
        ("局部视图", 3), ("斜视图", 3), ("向视图", 3), ("基本视图", 3), ("三视图", 3),
    ],
    "c5": [
        ("灯具", 3), ("照明", 3), ("光源", 3), ("电气", 3), ("绝缘", 3), ("触电", 3),
        ("安全", 2), ("电压", 2), ("电流", 2), ("防护", 2), ("IP等级", 3), ("镇流器", 3),
        ("灯座", 3), ("光通量", 3), ("照度", 3), ("温升", 3), ("接地", 3), ("短路", 3),
        ("熔断", 3), ("特低电压", 3), ("控制装置", 3), ("耐久性", 2), ("频闪", 3),
        ("显色", 3), ("发光", 2), ("白炽", 3), ("荧光灯", 3), ("LED", 2), ("冷色", 2),
    ],
    "c6": [
        ("CAD", 3), ("CAM", 3), ("CAE", 3), ("PDM", 3), ("PLM", 3), ("Pro/E", 3),
        ("Pro/Engineer", 3), ("SolidWorks", 3), ("UG", 3), ("AutoCAD", 3), ("Cimatron", 3),
        ("Rhino", 3), ("3dsMax", 3), ("软件", 2), ("建模", 3), ("参数化", 3), ("特征", 2),
        ("自顶向下", 3), ("自底向上", 3), ("装配建模", 3), ("变型设计", 3), ("三维造型", 3),
        ("工程绘图", 2), ("数控加工", 2), ("产品数据管理", 3), ("并行工程", 3), ("串行工程", 3),
        ("文档管理", 2), ("版本管理", 3), ("BOM", 3), ("Creo", 3), ("CAXA", 3), ("设计软件", 3),
        ("仿真", 3), ("数学模型", 3), ("逆向工程", 3), ("曲面连续", 3), ("几何关系", 2),
        ("草图", 2), ("抽壳", 3), ("阵列", 2),
    ],
}


def cut_dup_no(text, no):
    """PDF 里偶有同一题号出现两份（如 300 出现两次），第二份被并进了上一题的选项。
    只在「数字出现在文本中间且等于本题题号」时截断，避免误伤正常数值。"""
    m = re.search(r"(?<!\d)" + str(no) + r"\s*[.．、]\s*", text or "")
    if m and m.start() > 0:
        return text[:m.start()].rstrip()
    return text


def polish(r):
    """修题干末尾的排版残渣（判断题居多）。"""
    no = r["no"]
    r["stem"] = cut_dup_no(r["stem"].strip(), no)
    for o in r["options"]:
        o["text"] = cut_dup_no(o["text"], no)
        # 选项末尾多余的句点（PDF 每条选项后都带句号，去掉更像题目本身）
        if len(o["text"]) > 2:
            o["text"] = re.sub(r"[.．。]+$", "", o["text"]).strip()
    s = r["stem"].strip()
    if r["type"] == "judge":
        # 尾部多余的空括号，如 "熔融沉积快速成型工艺可以同时成型两种或以上材料（）。"
        s = re.sub(r"[（(]\s*[）)]\s*$", "", s).strip()
        # 尾部重复的答案标记，如 "...两地办学。 对" / "...外延意义范畴 错"
        tail = {"T": ("对", "正确", "√"), "F": ("错", "错误", "×")}.get(
            (r["answer"] or [None])[0], ())
        for t in tail:
            if s.endswith(t) and len(s) > len(t) + 3:
                s = s[: -len(t)].rstrip()
                break
        # 句子没句号就补一个，避免显示时句子断头
        if s and not s.endswith(("。", "？", "！", "?", "!", "：", ":")):
            s += "。"
    r["stem"] = s


def clean_for_match(s):
    return re.sub(r"\s+", "", s or "").upper()


def classify(stem, options):
    """关键词加权打分；返回 (chapter_id, 得分, 次高章节, 次高分)"""
    text = clean_for_match(stem + " " + " ".join(o["text"] for o in options))
    score = {}
    for cid, kws in KEYWORDS.items():
        s = 0
        for kw, w in kws:
            k = kw.upper()
            if k in text:
                s += w
        if s:
            score[cid] = s
    if not score:
        return "misc", 0, None, 0
    ranked = sorted(score.items(), key=lambda kv: (-kv[1], kv[0]))
    best, bscore = ranked[0]
    second = ranked[1] if len(ranked) > 1 else (None, 0)
    return best, bscore, second[0], second[1]


def norm(s):
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]", "", s or "").upper()


def dedupe_key(r):
    stem = norm(r["stem"])
    if r["type"] == "judge":
        return (stem, "judge", tuple(r["answer"]))
    texts = {o["key"]: norm(o["text"]) for o in r["options"]}
    return (stem, r["type"], tuple(sorted(texts.values())),
            tuple(sorted(texts.get(k, "") for k in r["answer"])))


ANSWER_TYPES = {"single", "multi", "judge"}


def main():
    recs = json.loads(RAW.read_text(encoding="utf-8"))
    by_no = {r["no"]: r for r in recs}

    # ---- 1. 补丁 ----
    for no, patch in PATCHES.items():
        r = by_no[no]
        if "options" in patch:
            letters = "ABCD"
            old = {o["key"]: o["text"] for o in r["options"]}
            opts = []
            for i, txt in enumerate(patch["options"]):
                key = letters[i]
                # 尽量沿用原文（补丁里没改的项保持原样）
                if key in old and old[key] == txt:
                    opts.append({"key": key, "text": txt})
                else:
                    opts.append({"key": key, "text": txt})
            r["options"] = opts
        if "answer" in patch:
            r["answer"] = list(patch["answer"])
            r["type"] = ("judge" if patch["answer"] in ("T", "F")
                         else ("multi" if len(patch["answer"]) > 1 else "single"))
        if "flag" in patch:
            r["flag"] = patch["flag"]
        if "note" in patch:
            r.setdefault("notes", []).append(patch["note"])

    # ---- 2. 去重 ----
    # 重复题不丢弃：标记 dupOf 后仍然输出，站点默认不抽取、设置里可开启
    seen = {}
    kept = []
    for r in recs:
        if r["type"] not in ANSWER_TYPES:
            continue
        k = dedupe_key(r)
        if k in seen:
            r["dupOf"] = f"q{seen[k]:04d}"
        else:
            seen[k] = r["no"]
        kept.append(r)

    # ---- 3. 题干收尾清理 + 章节归类 ----
    for r in kept:
        polish(r)
        cid, sc, c2, s2 = classify(r["stem"], r["options"])
        r["chapter"] = cid
        r["_score"] = sc
        r["_second"] = (c2, s2)

    # 未分类 / 低分题，按邻近已分类题的多数意见补齐（题库本身成块排列）
    nums = [r["no"] for r in kept]
    pos = {r["no"]: i for i, r in enumerate(kept)}
    changed = 0
    for r in kept:
        if r["chapter"] != "misc" and r["_score"] >= 2:
            continue
        i = pos[r["no"]]
        votes = Counter()
        for j in range(max(0, i - 8), min(len(kept), i + 9)):
            nb = kept[j]
            if nb is r or nb["chapter"] == "misc" or nb["_score"] < 2:
                continue
            votes[nb["chapter"]] += 1
        if votes:
            top, cnt = votes.most_common(1)[0]
            if cnt >= 5:  # 窗口内多数一致才采信
                r["chapter"] = top
                r["_smooth"] = True
                changed += 1
    print(f"按邻近题块补齐章节: {changed} 题")

    # ---- 4. 输出 data/questions.json ----
    qs = []
    for r in kept:
        qs.append({
            "id": f"q{r['no']:04d}",
            "no": r["no"],
            "page": r["page"],
            "type": r["type"],
            "chapter": r["chapter"],
            "stem": r["stem"],
            "options": [{"key": o["key"], "text": o["text"]} for o in r["options"]],
            "answer": r["answer"],
            "explain": "",
            "flag": r.get("flag", ""),
            "dupOf": r.get("dupOf"),
        })
    usable = [q for q in qs if not q["dupOf"] and q.get("flag") != "broken"]
    cnt = Counter(q["chapter"] for q in usable)
    chapter_rows = []
    for i, (cid, name) in enumerate(CHAPTERS):
        if cnt.get(cid):
            chapter_rows.append({"id": cid, "name": name, "order": len(chapter_rows) + 1,
                                 "count": cnt[cid]})
    data = {
        "meta": {"title": TITLE, "version": VERSION, "total": len(qs),
                 "activeTotal": len(usable)},
        "chapters": chapter_rows,
        "questions": qs,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")

    # ---- 统计 ----
    dup_n = sum(1 for q in qs if q["dupOf"])
    print(f"原始解析 {len(recs)} 题 -> 输出 {len(qs)} 题（其中重复题 {dup_n}，可用 {len(qs) - dup_n}）")
    print("题型:", dict(Counter(q["type"] for q in qs)))
    print("章节(可练):", dict(cnt))
    print("待复核(flag):", dict(Counter(q["flag"] or "ok" for q in qs)))

    write_xlsx(recs, by_no, kept)
    print(f"写出 {OUT_JSON}")
    print(f"写出 {OUT_XLSX}")


def write_xlsx(recs, by_no, kept):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    kept_nos = {r["no"] for r in kept}
    chname = {cid: name for cid, name in CHAPTERS}
    typ_name = {"single": "单选", "multi": "多选", "judge": "判断"}
    red = PatternFill("solid", fgColor="FFD6D6")
    yellow = PatternFill("solid", fgColor="FFF3CD")
    head_fill = PatternFill("solid", fgColor="DCE6F1")

    wb = Workbook()

    # --- 概览 ---
    ws = wb.active
    ws.title = "概览"
    rows = [
        ("题库", TITLE),
        ("版本", VERSION),
        ("原始题数", len(recs)),
        ("去重后题数", len(kept)),
        ("重复题（已标记 dupOf，默认不抽取）", len(recs) - len(kept)),
        ("", ""),
        ("题型", ""),
    ]
    for t, cnt in Counter(r["type"] for r in kept).items():
        rows.append((f"  {typ_name.get(t, t)}", cnt))
    rows.append(("", ""))
    rows.append(("章节", ""))
    for cid, name in CHAPTERS:
        rows.append((f"  {name}", sum(1 for r in kept if r["chapter"] == cid)))
    rows.append(("", ""))
    rows += [
        ("红色行=需人工确认（答案空缺/选项缺失/排版损坏）", ""),
        ("黄色行=已由我补答案或重建选项，请重点核对", ""),
        ("校对方法：直接在本表改答案/题干，改完把文件发我，我重新生成", ""),
    ]
    for i, (a, b) in enumerate(rows, start=1):
        ws.cell(row=i, column=1, value=a)
        ws.cell(row=i, column=2, value=b)
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 50
    ws["A1"].font = Font(bold=True, size=13)

    # --- 全部题目 ---
    ws = wb.create_sheet("全部题目")
    headers = ["PDF题号", "页码", "题型", "章节", "题干", "A", "B", "C", "D", "答案", "备注", "重复于"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center")

    def note_of(r):
        parts = list(r.get("notes", []))
        if r.get("issues"):
            parts = list(r["issues"]) + parts
        if r.get("_smooth"):
            parts.append("章节按邻近题块推断")
        return "；".join(parts)

    for r in recs:
        opt = {o["key"]: o["text"] for o in r["options"]}
        flag = r.get("flag", "")
        row = [
            r["no"], r["page"], typ_name.get(r["type"], r["type"] or "?"),
            chname.get(r.get("chapter", ""), "") if r["no"] in kept_nos else "",
            r["stem"],
            opt.get("A", ""), opt.get("B", ""), opt.get("C", ""), opt.get("D", ""),
            "".join(r["answer"]) if r["type"] != "judge" else ("正确" if r["answer"] == ["T"] else "错误" if r["answer"] == ["F"] else ""),
            note_of(r),
            r.get("dupOf") or "",
        ]
        ws.append(row)
        ri = ws.max_row
        fill = None
        if flag in ("review", "ai-answer"):
            fill = red
        elif r.get("issues") or (r["no"] not in kept_nos):
            fill = yellow
        if fill:
            for cell in ws[ri]:
                cell.fill = fill
        for col in range(1, len(headers) + 1):
            ws.cell(row=ri, column=col).alignment = Alignment(vertical="top", wrap_text=(col == 5))

    widths = [8, 6, 6, 16, 60, 22, 22, 22, 22, 8, 42, 8]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    # --- 待确认 ---
    ws = wb.create_sheet("待确认")
    ws.append(["PDF题号", "页码", "题型", "问题", "题干", "选项", "当前答案", "建议"])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = head_fill
    for r in recs:
        flag = r.get("flag", "")
        if flag not in ("review", "ai-answer", "broken"):
            continue
        opts = " / ".join(f"{o['key']}.{o['text']}" for o in r["options"])
        ans = "".join(r["answer"]) if r["type"] != "judge" else ("正确" if r["answer"] == ["T"] else "错误" if r["answer"] == ["F"] else "")
        if flag == "ai-answer":
            problem = "原PDF答案为空，已按标准答案补"
            advice = "请确认答案是否为你考试使用的标准答案"
        elif flag == "broken":
            problem = "题目损坏，无法还原，站点已永久排除"
            advice = "建议从题库中删除（在 全部题目 表里把整行删掉不影响其他题）"
        else:
            problem = "选项缺失/排版损坏，已按原文重建"
            advice = "请对照原PDF核对选项与答案"
        ws.append([r["no"], r["page"], typ_name.get(r["type"], "?"), problem, r["stem"], opts, ans, advice])
        ri = ws.max_row
        for cell in ws[ri]:
            cell.fill = PatternFill("solid", fgColor="E8E8E8") if flag == "broken" else (red if flag == "ai-answer" else yellow)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for i, w in enumerate([8, 6, 6, 24, 50, 80, 10, 34], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    # --- 重复题 ---
    ws = wb.create_sheet("重复题")
    ws.append(["保留题号", "重复题号", "题型", "题干"])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = head_fill
    for r in recs:
        if r.get("dupOf"):
            ws.append([r["dupOf"], r["no"], typ_name.get(r["type"], "?"), r["stem"]])
    for i, w in enumerate([10, 10, 6, 90], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_XLSX)


if __name__ == "__main__":
    main()
