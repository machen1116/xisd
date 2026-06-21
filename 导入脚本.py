#!/usr/bin/env python3
"""
考点 Excel → index.html 导入脚本（3列格式，支持模式选择）
使用方法：
  1. Excel 格式：章节 | 正面 | 反面（前三列，列名任意）
  2. 运行：python3 导入脚本.py
  3. 选择导入模式：重点复习 / 学习模式
  4. 可反复运行，每次只追加新内容（自动去重）

命令行参数：
  python3 导入脚本.py --mode review   # 直接导入到重点复习
  python3 导入脚本.py --mode study    # 直接导入到学习模式
  python3 导入脚本.py --file 其他文件.xlsx  # 指定 Excel 文件
"""

import openpyxl
import re
import sys
import os
import argparse
import json as json_mod

EXCEL_FILE = "考点导入模板.xlsx"
HTML_FILE = "index.html"

MODE_MAP = {
    "重点复习": "review",
    "学习模式": "study",
    "复习": "review",
    "学习": "study",
    "review": "review",
    "study": "study",
}

VAR_MAP = {
    "review": "chapterData",
    "study": "studyData",
}


def load_excel_3col(path):
    """
    读取 Excel 前三列，格式：章节 | 正面 | 反面
    列名任意，只取前三列
    返回：{chapter: [cards...]}
    """
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    data = {}  # {chapter: [cards]}
    errors = []
    count = 0

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        chapter = row[0]
        front = row[1]
        back = row[2]

        if not chapter or not front or not back:
            continue

        chapter = str(chapter).strip()
        front = str(front).strip()
        back = str(back).strip()

        if not front or not back:
            continue

        card = {"front": front, "back": back}
        data.setdefault(chapter, []).append(card)
        count += 1

    print(f"  📊 共读取 {count} 张卡片，涉及 {len(data)} 个章节")
    return data, errors


def choose_mode():
    """让用户选择导入模式"""
    print("\n请选择导入目标模式：")
    print("  1) 重点复习（考前复习重点，对应已有内容）")
    print("  2) 学习模式（详细学习，对应后续新增内容）")
    print("  3) 退出")

    while True:
        choice = input("请输入数字 (1/2/3): ").strip()
        if choice == "1":
            return "review"
        elif choice == "2":
            return "study"
        elif choice == "3":
            print("已取消导入")
            sys.exit(0)
        else:
            print("请输入 1、2 或 3")


def build_name_to_key_map(html):
    """
    从 HTML 中解析章节显示名 → key 的映射。
    匹配：<div class="chapter-card ch-XXX" onclick="goToModePage('key')"> ... <h3>显示名</h3>
    """
    # 正则：捕获 class 里的 ch-KEY，onclick 里的 key，以及 <h3> 里的显示名
    pattern = r'class="chapter-card ch-(\w+)"[^>]*onclick="goToModePage\(\'(\w+)\'\)"[\s\S]*?<h3>(.*?)</h3>'
    matches = re.findall(pattern, html)
    mapping = {}
    for (ch_class, key, display_name) in matches:
        display_name = display_name.strip()
        mapping[display_name] = key
    return mapping


def parse_js_var(html, var_name):
    """
    从 HTML 中解析 JS 变量（对象形式），用 JSON 解析。
    返回解析到的 dict，解析失败返回 {}。
    """
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

    js_block = html[start:end+1]  # 包含 {}
    # 去掉尾逗号（JSON 不允许）
    cleaned = re.sub(r',\s*([}\]])', r'\1', js_block)
    # 现在是有效的 JSON，用 json.loads() 解析
    try:
        data = json_mod.loads(cleaned)
        return data
    except json_mod.JSONDecodeError as e:
        print(f"  ⚠️ JSON 解析失败：{e}")
        # fallback: 返回空 dict
        return {}


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
    parser = argparse.ArgumentParser(description="考点 Excel → index.html 导入脚本")
    parser.add_argument("--mode", choices=["review", "study"], help="导入模式：review=重点复习，study=学习模式")
    parser.add_argument("--file", default=EXCEL_FILE, help=f"Excel 文件名（默认：{EXCEL_FILE}）")
    args = parser.parse_args()

    excel_file = args.file

    # 检查 Excel 文件是否存在
    if not os.path.exists(excel_file):
        print(f"❌ 找不到 {excel_file}，请把 Excel 文件放在同一目录下")
        sys.exit(1)

    print(f"📖 读取 {excel_file}（3列格式：章节 | 正面 | 反面）...")
    new_data, errors = load_excel_3col(excel_file)

    if errors:
        print("⚠️  发现以下问题：")
        for e in errors:
            print(f"  - {e}")

    if not new_data:
        print("❌ Excel 中没有有效数据，请检查填写内容")
        sys.exit(1)

    # 读取 HTML（用于解析映射表和已有数据）
    print(f"\n📝 读取 {HTML_FILE} ...")
    with open(HTML_FILE, "r", encoding="utf-8") as f:
        html = f.read()

    # 解析映射表：显示名 → key
    name_to_key = build_name_to_key_map(html)
    print(f"  📋 找到 {len(name_to_key)} 个章节映射")

    # 把 Excel 里的显示名转换成 key
    new_data_key = {}
    for display_name, cards in new_data.items():
        if display_name in name_to_key:
            key = name_to_key[display_name]
            new_data_key[key] = cards
        else:
            print(f"  ⚠️  警告：章节「{display_name}」在网页中找不到对应，已跳过")

    if not new_data_key:
        print("❌ 没有有效章节可导入，请检查 Excel 中的章节名称是否准确")
        sys.exit(1)

    new_data = new_data_key

    # 选择模式（命令行参数 或 交互选择）
    if args.mode:
        mode = args.mode
        print(f"📌 命令行指定模式：{mode}")
    else:
        mode = choose_mode()

    var_name = VAR_MAP[mode]
    mode_label = "重点复习" if mode == "review" else "学习模式"

    # 读取已有数据
    existing = parse_js_var(html, var_name)
    if existing:
        print(f"  ✅ 已找到现有【{mode_label}】数据：{len(existing)} 个章节")
    else:
        print(f"  ℹ️  未找到现有【{mode_label}】数据，将新建")

    # 合并数据
    merged, added = merge_data(existing, new_data)
    print(f"  ➕ 新增 {added} 张卡片（已去重）")

    if added == 0:
        print("  ℹ️  没有新内容需要导入")
        sys.exit(0)

    # 写回 HTML
    new_js = build_chapterdata_js(merged, var_name)
    html = replace_js_var_in_html(html, var_name, new_js)

    print(f"\n💾 写入 {HTML_FILE} ...")
    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n✅ 导入完成！")
    print(f"   模式：{mode_label}")
    print(f"   新增：{added} 张")
    print(f"   请刷新浏览器查看效果")
    print(f"   在线版：https://machen1116.github.io/xisd/")


if __name__ == "__main__":
    main()
