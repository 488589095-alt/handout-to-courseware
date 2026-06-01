"""
parse_handout.py — 解析讲义 docx → handout.json

输入：讲义 docx 文件
输出：符合 references/handout_schema.md 的 JSON

用法：
  python3 parse_handout.py 第1讲分解质因数.docx -o output/handout.json
"""

import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document


# 正则：题目识别
ITEM_PATTERNS = [
    # 例题1【BYGF】题干...
    (re.compile(r"^例题\s*(\d+)\s*(?:【(.+?)】)?\s*(.*)"), "例题"),
    (re.compile(r"^例\s+(\d+)\s*(?:【(.+?)】)?\s*(.*)"), "例题"),
    # 练习1【BYGF】题干...
    (re.compile(r"^练习\s*(\d+)\s*(?:【(.+?)】)?\s*(.*)"), "练习"),
    (re.compile(r"^练\s+(\d+)\s*(?:【(.+?)】)?\s*(.*)"), "练习"),
]


def match_item_start(text: str):
    """匹配题目起始，返回 (kind, num, tag, body_first_line) 或 None。"""
    for pattern, kind in ITEM_PATTERNS:
        m = pattern.match(text)
        if m:
            num = int(m.group(1))
            tag = m.group(2) or None
            body = m.group(3).strip()
            return kind, num, tag, body
    return None


def is_objective_line(text: str):
    """识别课程目标行，返回 (key, value) 或 None。"""
    keys = ["知识技能", "数学能力", "思想方法", "情感态度"]
    for k in keys:
        if text.startswith(k):
            # 去掉冒号后取剩余部分
            rest = text[len(k):].lstrip("：:").strip()
            return k, rest
    return None


def parse_handout(docx_path: Path) -> dict:
    doc = Document(docx_path)

    handout = {
        "title": "",
        "lecture_num": "",
        "objectives": {},
        "knowledge_ladder": None,
        "modules": [],
        "innovation": None,
    }

    current_module = None
    current_item = None
    section = "preamble"  # preamble / objectives / knowledge_ladder / module / innovation

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        style = p.style.name if p.style else ""

        # ── 首段标题 ──
        if not handout["title"] and "讲" in text and len(text) < 30:
            # "第一讲 分解质因数" → lecture_num="第一讲" or "第1讲"
            handout["title"] = text
            m = re.match(r"^(第[一二三四五六七八九十百\d]+讲)\s+(.+)", text)
            if m:
                handout["lecture_num"] = m.group(1)
            continue

        # ── Heading 识别 ──
        if "Heading 1" in style:
            if text == "知识阶梯":
                section = "knowledge_ladder"
                handout["knowledge_ladder"] = ""
                continue
            elif text == "课程目标":
                section = "objectives"
                continue
            else:
                # 模块标题
                section = "module"
                current_module = {
                    "module_num": len(handout["modules"]) + 1,
                    "title": text,
                    "items": [],
                }
                handout["modules"].append(current_module)
                current_item = None
                continue

        if "Heading 2" in style:
            if text == "创新挑战":
                section = "innovation"
                handout["innovation"] = {"kind": "创新挑战", "num": 1, "tag": None, "body": ""}
                continue

        # ── 内容处理 ──
        if section == "knowledge_ladder":
            handout["knowledge_ladder"] = (handout["knowledge_ladder"] or "") + "\n" + text
            continue

        if section == "objectives":
            kv = is_objective_line(text)
            if kv:
                handout["objectives"][kv[0]] = kv[1]
            continue

        if section == "module" and current_module:
            # 看是不是新题
            m = match_item_start(text)
            if m:
                kind, num, tag, body_first = m
                current_item = {
                    "kind": kind,
                    "num": num,
                    "tag": tag,
                    "body": body_first,
                    "solution": None,
                }
                current_module["items"].append(current_item)
            else:
                # 继续上个题的题干
                if current_item is not None:
                    current_item["body"] += "\n" + text
            continue

        if section == "innovation" and handout["innovation"] is not None:
            # 创新挑战的 body
            # 可能首段是 "1【tag】题干"，也可能直接是题干
            m = match_item_start(text)
            if m:
                kind, num, tag, body_first = m
                handout["innovation"]["num"] = num
                handout["innovation"]["tag"] = tag
                handout["innovation"]["body"] = body_first
            elif text.startswith("1") and "【" in text:
                # 形如 "1【2020 • 专项】..."
                m2 = re.match(r"^\d+\s*【(.+?)】\s*(.*)", text)
                if m2:
                    handout["innovation"]["tag"] = m2.group(1)
                    handout["innovation"]["body"] = m2.group(2)
                else:
                    handout["innovation"]["body"] += ("\n" if handout["innovation"]["body"] else "") + text
            else:
                handout["innovation"]["body"] += ("\n" if handout["innovation"]["body"] else "") + text
            continue

    # 收尾清理
    if handout["knowledge_ladder"]:
        handout["knowledge_ladder"] = handout["knowledge_ladder"].strip()

    return handout


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="讲义 docx 文件")
    parser.add_argument("-o", "--output", required=True, help="输出 JSON 文件")
    args = parser.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        sys.exit(f"❌ 讲义不存在：{in_path}")

    handout = parse_handout(in_path)

    # 简单 sanity check
    if not handout["title"]:
        print(f"⚠️  没识别到标题", file=sys.stderr)
    if not handout["modules"]:
        sys.exit(f"❌ 没识别到任何模块，请检查讲义 Heading 1 是否正确")

    # 写文件
    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(handout, ensure_ascii=False, indent=2), encoding="utf-8")

    # 报告
    total_items = sum(len(m["items"]) for m in handout["modules"])
    print(f"✅ 解析成功：{handout['title']}")
    print(f"   模块：{len(handout['modules'])} 个")
    print(f"   题目：{total_items} 道（{sum(1 for m in handout['modules'] for i in m['items'] if i['kind']=='例题')} 例 + {sum(1 for m in handout['modules'] for i in m['items'] if i['kind']=='练习')} 练）")
    print(f"   创新挑战：{'✅' if handout['innovation'] else '❌'}")
    print(f"   输出：{out_path}")


if __name__ == "__main__":
    main()
