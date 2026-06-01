"""
parse_teaching_design.py — 解析教学设计稿 docx → teaching_design.json

输入：教学设计稿 docx
输出：符合 references/teaching_design_schema.md 的 JSON

用法：
  python3 parse_teaching_design.py 分解质因数110分钟教学设计.docx -o output/teaching_design.json
"""

import argparse
import json
import re
import sys
from pathlib import Path

from docx import Document


# Activity 起始关键词
ACTIVITY_KEYWORDS = {
    "情境创设": "情境创设",
    "探究活动": "探究活动",
    "技能学习": "技能学习",
    "例题解析": "例题解析",
    "学生练习": "学生练习",
    "故事推进": "故事推进",
    "核心概念讲解": "核心概念讲解",
    "方法总结": "方法总结",
    "方法迁移": "方法迁移",
    "题目分析": "题目分析",
    "破解过程": "破解过程",
    "尝试探索": "尝试探索",
    "思路引导": "思路引导",
}


def extract_duration(text: str):
    """从文本中提取 '（X分钟）' 的数字。"""
    m = re.search(r"[（(](\d+)\s*分钟[）)]", text)
    return int(m.group(1)) if m else None


def extract_references(text: str):
    """从 activity 文本中提取引用的讲义题号。"""
    refs = []
    # 例题N / 例N
    for m in re.finditer(r"例题\s*(\d+)|例\s+(\d+)", text):
        num = int(m.group(1) or m.group(2))
        refs.append({"kind": "例题", "num": num})
    # 练习N / 练N
    for m in re.finditer(r"练习\s*(\d+)|练\s+(\d+)", text):
        num = int(m.group(1) or m.group(2))
        refs.append({"kind": "练习", "num": num})
    # 去重
    seen = set()
    out = []
    for r in refs:
        k = (r["kind"], r["num"])
        if k not in seen:
            seen.add(k)
            out.append(r)
    return out


def detect_activity_start(text: str):
    """检测段落是否是 activity 起始，返回 activity_type 或 None。"""
    # 严格匹配："XXX (X分钟)：" 或 "XXX："
    for keyword, atype in ACTIVITY_KEYWORDS.items():
        if text.startswith(keyword) and ("：" in text[:30] or "(" in text[:30] or "（" in text[:30]):
            return atype
    return None


def classify_segment(heading_text: str):
    """根据 Heading 文本判断 segment 类型。"""
    # 故事线 / 教学亮点 / 教学准备 / 预期效果 = meta
    meta_keywords = ["亮点", "教学准备", "预期教学效果", "故事线"]
    for k in meta_keywords:
        if k in heading_text:
            return "meta"
    # 总结 / 拓展 优先级高于 模块（避免"总结与拓展"被误判为模块）
    if "总结" in heading_text or "拓展" in heading_text:
        return "总结"
    if "导入" in heading_text:
        return "导入"
    if "模块" in heading_text:
        return "模块"
    return "其他"


def parse_teaching_design(docx_path: Path) -> dict:
    doc = Document(docx_path)

    design = {
        "title": "",
        "total_duration_min": None,
        "story_line": None,
        "segments": [],
        "meta": {
            "time_allocation": "",
            "story_recap": [],
            "key_points": [],
            "extension_challenges": "",
            "teaching_highlights": [],
        },
    }

    current_segment = None
    current_activity = None
    current_meta_section = None  # 当前在哪个 meta 区块

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        style = p.style.name if p.style else ""

        # ── Heading 1：总标题 ──
        if "Heading 1" in style:
            if not design["title"]:
                design["title"] = text
                # 从标题提取时长，如 "《分解质因数》110分钟教学设计"
                m = re.search(r"(\d+)\s*分钟", text)
                if m:
                    design["total_duration_min"] = int(m.group(1))
            continue

        # ── Heading 2：segment 或 meta 区 ──
        if "Heading 2" in style:
            seg_type = classify_segment(text)

            if seg_type == "meta":
                # 进入 meta 区
                current_segment = None
                current_activity = None
                # 特判：故事线 → 单独标记 story_line
                if "故事线" in text:
                    current_meta_section = "story_line"
                    if design["story_line"] is None:
                        # 从 heading 提取标题，如"故事线：密码破解者的冒险"
                        story_title = text.split("：", 1)[-1].strip() if "：" in text else text
                        design["story_line"] = {
                            "title": story_title,
                            "background": "",
                            "student_role": "",
                        }
                else:
                    current_meta_section = text
                continue

            # 否则是 segment
            seg_num_match = re.match(r"^([一二三四五六七八九十]+)、", text)
            seg_num = seg_num_match.group(1) if seg_num_match else ""

            # 提取模块标题（如"二、模块一：分解质因数的方法（25分钟）"→ "分解质因数的方法"）
            module_title = text
            m = re.search(r"模块[一二三四五六七八九十]+：(.+?)(?:（|$)", text)
            if m:
                module_title = m.group(1).strip()
            elif "：" in text:
                # 如"一、导入环节（15分钟）" 取冒号后
                module_title = text.split("：", 1)[-1].strip() if "：" in text else text

            current_segment = {
                "segment_num": seg_num,
                "segment_type": seg_type,
                "module_title": module_title,
                "duration_min": extract_duration(text),
                "subtitle": None,
                "activities": [],
            }
            design["segments"].append(current_segment)
            current_activity = None
            current_meta_section = None
            continue

        # ── Heading 3：meta sub-sections ──
        if "Heading 3" in style:
            if "时间分配" in text:
                current_meta_section = "time_allocation"
            elif "故事线" in text and "回顾" in text:
                current_meta_section = "story_recap"
            elif "知识点" in text and ("梳理" in text or "回顾" in text):
                current_meta_section = "key_points"
            elif "拓展" in text:
                current_meta_section = "extension_challenges"
            elif "亮点" in text or "亮点" in (current_meta_section or ""):
                # 设计亮点的子节
                pass
            continue

        # ── Normal 段落 ──
        # 在 meta 区
        if current_meta_section:
            if current_meta_section == "time_allocation":
                design["meta"]["time_allocation"] += text + "\n"
            elif current_meta_section == "story_recap":
                design["meta"]["story_recap"].append(text)
            elif current_meta_section == "key_points":
                design["meta"]["key_points"].append(text)
            elif current_meta_section == "extension_challenges":
                design["meta"]["extension_challenges"] += text + "\n"
            elif current_meta_section == "story_line":
                # 故事线段落
                if design["story_line"] is None:
                    design["story_line"] = {
                        "title": "",
                        "background": "",
                        "student_role": "",
                    }
                # 第一段含 "背景设定" → background
                if "背景设定" in text or "课程背景" in text:
                    design["story_line"]["background"] = text
                elif not design["story_line"]["background"]:
                    design["story_line"]["background"] = text
                # 推断 student_role
                if "扮演" in text or "化身" in text or "成为" in text:
                    m = re.search(r"(?:扮演|化身|成为)[为成]?[“”\"]?(.+?)[“”\"]?[，。\s]", text)
                    if m:
                        design["story_line"]["student_role"] = m.group(1)
            else:
                # 设计亮点等
                design["meta"]["teaching_highlights"].append(text)
            continue

        # 在 segment 内
        if current_segment is not None:
            # 检查是否是子标题（如"工具修炼：短除法秘籍"）
            if not current_segment["subtitle"] and not current_activity and ("：" in text or "秘籍" in text or "修炼" in text or "锻造" in text or "魔法" in text or "挑战" in text or "破译" in text):
                # 极简启发：第一个非 activity 段落是 subtitle
                # （注意：activity 关键词优先级更高）
                if not detect_activity_start(text):
                    current_segment["subtitle"] = text
                    continue

            # 检查是否是 activity 起始
            atype = detect_activity_start(text)
            if atype:
                current_activity = {
                    "activity_type": atype,
                    "duration_min": extract_duration(text),
                    "raw_text": text,
                    "bullet_items": [],
                    "references": [],
                }
                current_segment["activities"].append(current_activity)
                continue

            # 否则是当前 activity 的内容
            if current_activity is not None:
                current_activity["raw_text"] += "\n" + text
                # bullet item 识别
                if re.match(r"^\d+[\.\、]", text) or text.startswith("◦") or text.startswith("•"):
                    current_activity["bullet_items"].append(text)
            else:
                # 在 segment 但还没遇到 activity → 算 subtitle 之后的引导段
                # 创建一个"开场白"虚拟 activity
                current_activity = {
                    "activity_type": "情境创设",  # 默认归入情境创设
                    "duration_min": None,
                    "raw_text": text,
                    "bullet_items": [],
                    "references": [],
                }
                current_segment["activities"].append(current_activity)

    # 后处理：每个 activity 提取 references
    for seg in design["segments"]:
        for act in seg["activities"]:
            act["references"] = extract_references(act["raw_text"])

    # 清理空字符串
    design["meta"]["time_allocation"] = design["meta"]["time_allocation"].strip()
    design["meta"]["extension_challenges"] = design["meta"]["extension_challenges"].strip()

    return design


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="教学设计稿 docx 文件")
    parser.add_argument("-o", "--output", required=True, help="输出 JSON 文件")
    args = parser.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        sys.exit(f"❌ 教学设计稿不存在：{in_path}")

    design = parse_teaching_design(in_path)

    # sanity check
    if not design["title"]:
        print("⚠️  没识别到标题", file=sys.stderr)
    if not design["segments"]:
        sys.exit("❌ 没识别到任何 segment")

    # 写文件
    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(design, ensure_ascii=False, indent=2), encoding="utf-8")

    # 报告
    print(f"✅ 解析成功：{design['title']}")
    print(f"   总时长：{design.get('total_duration_min')} 分钟")
    print(f"   故事线：{'✅' if design['story_line'] else '❌'}")
    print(f"   segments：{len(design['segments'])} 个")
    for seg in design["segments"]:
        refs = sum(len(a["references"]) for a in seg["activities"])
        print(f"     - [{seg['segment_type']}] {seg['module_title']}：{len(seg['activities'])} 活动, {refs} 个题目引用")
    print(f"   输出：{out_path}")


if __name__ == "__main__":
    main()
