#!/usr/bin/env python3
"""
考点 Excel → index.html 导入脚本（支持复习/学习双模式）
使用方法：
  1. 填写"考点导入模板.xlsx"（注意「模式」列）
  2. 运行本脚本：python3 导入脚本.py
  3. 脚本根据「模式」列自动写入对应数据集
  4. 可反复运行，每次只追加新内容（自动去重）

模式列填写说明：
  - 填「复习」或留空 → 写入复习模式数据（已有内容）
  - 填「学习」→ 写入学习模式数据（后续新增内容）
"""

import openpyxl
import re
import sys
import json

EXCEL_FILE = "考点导入模板.xlsx"
HTML_FILE = "index.html"

MODE_MAP = {
    "复习": "review",
    "学习": "study",
    "review": "review",
    "study": "study",
}


def load_excel(path):
    """读取 Excel，按模式分组返回 {mode_key: {章节: [cards...]}}"""
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    data = {}  # {mode_key: {chapter: [cards]}}
    errors = []

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        chapter, mode, qtype, front, back = row[0], row[1], row[2], row[3], row[4]
        if not chapter or not front or not back:
            continue
        chapter = str(chapter).strip()
        mode = str(mode).strip() if mode else "复习"
        front = str(front).strip()
        back = str(back).strip()

        mode_key = MODE_MAP.get(mode, "review")
        card = {"front": front, "back": back}
        data.setdefault(mode_key, {}).setdefault(chapter, []).append(card)

    return data, errors


def parse_js_var(html, var_name):
    """
    从 HTML 中解析 JS 变量（对象形式）。
    返回解析到的 dict，解析失败返回 {}。
    """
    # 匹配 var xxx = { ... };
    pattern = r"var\s+" + re.escape(var_name) + r"\s*=\s*\{(.*?)\}\s*;"
    # 但上面非贪婪匹配会过早结束，改用：找到变量起始，然后手工解析大括号
    marker = f"var {var_name} = "
    idx = html.find(marker)
    if idx == -1:
        return {}
    start = idx + len(marker)
    # 找到匹配的最外层 {}
    brace_count = 0
    in_string = False
    escape_next = False
    end = None
    i = start
    while i < len(html):
        ch = html[i]
        if escape_next:
            escape_next = False
            i += 1
            continue
        if ch == '\\':
            escape_next = True
            i += 1
            continue
        if ch == '"' or ch == "'":
            if not in_string:
                in_string = ch
            elif in_string == ch:
                in_string = False
            i += 1
            continue
        if in_string:
            i += 1
            continue
        if ch == '{':
            brace_count += 1
        elif ch == '}':
            brace_count -= 1
            if brace_count == 0:
                end = i
                break
        i += 1

    if end is None:
        return {}

    js_block = html[start+1:end].strip()  # 去掉外层 {}
    # 现在解析每个章节键
    chapters = {}
    # 用正则找 "chapterKey": [ ... ]
    # 章节键可能是 "dalun" 或 "ch1" 等
    ch_pattern = r'"(\w+)"\s*:\s*\[(.*?)\]'
    for m in re.finditer(ch_pattern, js_block, re.DOTALL):
        ch_key = m.group(1)
        cards_str = m.group(2)
        cards = []
        # 解析每张卡片
        for cm in re.finditer(r'\{\s*"id"\s*:\s*"[^"]*",\s*"front"\s*:\s*"((?:[^"\\]|\\.)*)",\s*"back"\s*:\s*"((?:[^"\\]|\\.)*)"(?:\s*,\s*"tip"\s*:\s*"((?:[^"\\]|\\.)*)")?\s*\}', cards_str, re.DOTALL):
            f = cm.group(1).replace('\\"', '"').replace('\\\\', '\\')
            b = cm.group(2).replace('\\"', '"').replace('\\\\', '\\')
            tip = cm.group(4)
            if tip:
                tip = tip.replace('\\"', '"').replace('\\\\', '\\')
            card = {"front": f, "back": b}
            if tip:
                card["tip"] = tip
            cards.append(card)
        chapters[ch_key] = cards

    return chapters


def merge_data(existing, new_cards_by_chapter):
    """将新卡片合并到已有数据，去重"""
    merged = dict(existing)
    added = 0
    for chapter, new_cards in new_cards_by_chapter.items():
        existing_cards = merged.get(chapter, [])
        existing_keys = set((c["front"], c["back"]) for c in existing_cards)
        for card in new_cards:
            k = (card["front"], card["back"])
            if k not in existing_keys:
                existing_cards.append(card)
                existing_keys.add(k)
                added += 1
        if chapter not in merged:
            merged[chapter] = existing_cards
    return merged, added


def build_chapterdata_js(data, var_name):
    """构建 JS 变量赋值语句"""
    lines = []
    lines.append(f"        var {var_name} = {{")
    chapters = list(data.items())
    for i, (ch_key, cards) in enumerate(chapters):
        # 找章节显示名
        lines.append(f'            "{ch_key}": [')
        for card in cards:
            front_escaped = card["front"].replace('\\', '\\\\').replace('"', '\\"')
            back_escaped = card["back"].replace('\\', '\\\\').replace('"', '\\"')
            tip_str = ""
            if "tip" in card:
                tip_escaped = card["tip"].replace('\\', '\\\\').replace('"', '\\"')
                tip_str = f', "tip": "{tip_escaped}"'
            lines.append(f'                {{ "id": "", "front": "{front_escaped}", "back": "{back_escaped}"{tip_str} }},')
        lines.append("            ]")
        if i < len(chapters) - 1:
            lines.append("            },")
        else:
            lines.append("            }")
    lines.append("        };")
    return "\n".join(lines)


def replace_js_var_in_html(html, var_name, new_js_block):
    """替换 HTML 中 var var_name = ... 的内容"""
    marker = f"var {var_name} = "
    idx = html.find(marker)
    if idx == -1:
        # 变量不存在，需要插入（追加到 </script> 前）
        return html.replace("</script>", new_js_block + "\n    </script>", 1)

    start = idx + len(marker)
    brace_count = 0
    in_string = False
    escape_next = False
    end = None
    i = start
    while i < len(html):
        ch = html[i]
        if escape_next:
            escape_next = False
            i += 1
            continue
        if ch == '\\':
            escape_next = True
            i += 1
            continue
        if ch == '"' or ch == "'":
            if not in_string:
                in_string = ch
            elif in_string == ch:
                in_string = False
            i += 1
            continue
        if in_string:
            i += 1
            continue
        if ch == '{':
            brace_count += 1
        elif ch == '}':
            brace_count -= 1
            if brace_count == 0:
                end = i
                break
        i += 1

    if end is None:
        print(f"❌ 无法解析变量 {var_name} 的结束位置")
        sys.exit(1)

    new_block = f"var {var_name} = {new_js_block}"
    html = html[:idx] + new_block + html[end+1:]
    return html


def main():
    print(f"📖 读取 {EXCEL_FILE} ...")
    new_data, errors = load_excel(EXCEL_FILE)

    if errors:
        print("⚠️  发现以下问题：")
        for e in errors:
            print(f"  - {e}")

    if not new_data:
        print("❌ Excel 中没有有效数据，请检查填写内容")
        sys.exit(1)

    total_added = 0
    total_review = 0
    total_study = 0

    print(f"\n📝 读取 {HTML_FILE} 中已有内容 ...")
    with open(HTML_FILE, "r", encoding="utf-8") as f:
        html = f.read()

    # 处理复习模式
    if "review" in new_data:
        print("\n🔵 处理【复习模式】数据 ...")
        existing_review = parse_js_var(html, "chapterData")
        if existing_review:
            print(f"  ✅ 已找到现有复习数据：{len(existing_review)} 个章节")
        merged_review, added = merge_data(existing_review, new_data["review"])
        total_added += added
        total_review = added
        print(f"  ➕ 复习模式新增 {added} 张卡片")
        # 写回 HTML
        new_review_js = build_chapterdata_js(merged_review, "chapterData")
        html = replace_js_var_in_html(html, "chapterData", new_review_js)
        print(f"  ✅ 复习模式数据已更新（共 {len(merged_review)} 个章节）")

    # 处理学习模式
    if "study" in new_data:
        print("\n🟦 处理【学习模式】数据 ...")
        existing_study = parse_js_var(html, "studyData")
        if existing_study:
            print(f"  ✅ 已找到现有学习数据：{len(existing_study)} 个章节")
        merged_study, added = merge_data(existing_study, new_data["study"])
        total_added += added
        total_study = added
        print(f"  ➕ 学习模式新增 {added} 张卡片")
        # 写回 HTML
        new_study_js = build_chapterdata_js(merged_study, "studyData")
        html = replace_js_var_in_html(html, "studyData", new_study_js)
        print(f"  ✅ 学习模式数据已更新（共 {len(merged_study)} 个章节）")

    # 写回文件
    print(f"\n💾 写入 {HTML_FILE} ...")
    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n✅ 导入完成！")
    print(f"   复习模式新增：{total_review} 张")
    print(f"   学习模式新增：{total_study} 张")
    print(f"   共计新增：{total_added} 张")
    print(f"   请刷新浏览器查看效果")


if __name__ == "__main__":
    main()
