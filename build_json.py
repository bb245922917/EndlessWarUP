#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公告表格 -> announcement.json

用法:
    python build_json.py            读取 announcements.xlsx，生成 announcement.json
    python build_json.py --init     生成一份空的 announcements.xlsx 模板（不会覆盖已有文件）
    python build_json.py --check    只检查表格内容是否合法，不写文件

表格字段（第一行是表头，从第二行开始是数据）:
    id          公告唯一 ID，用于记录"已读"，不要重复、不要改
    title       标题
    date        日期，格式 2026-09-09
    tag         标签：更新 / 活动 / 维护 / 补偿 / 公告
    content     正文，支持换行（单元格内 Alt+Enter 换行）
    image       图片地址，可留空
    important   TRUE/FALSE，是否强制弹窗（FALSE 只在公告列表里显示）
    min_code    最低版本号(versionCode)才显示，0 = 不限
    max_code    最高版本号，0 = 不限
    lang        语言：zh / en
    enabled     TRUE/FALSE，FALSE 表示不下发
"""

import json
import os
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "announcements.xlsx")
JSON_OUT = os.path.join(HERE, "announcement.json")

FIELDS = [
    "id", "title", "date", "tag", "content", "image",
    "important", "min_code", "max_code", "lang", "enabled",
]

SAMPLE_ROWS = [
    {
        "id": "A20260909001",
        "title": "新版本 0.1.6.4 更新内容",
        "date": "2026-09-09",
        "tag": "更新",
        "content": "1. 优化了低端机型的帧率表现\r\n2. 新增战斗中随机奖励热气球\r\n3. 修复若干已知问题",
        "image": "",
        "important": True,
        "min_code": 185,
        "max_code": 0,
        "lang": "zh",
        "enabled": True,
    },
    {
        "id": "A20260909002",
        "title": "周末双倍活动开启",
        "date": "2026-09-12",
        "tag": "活动",
        "content": "本周六、周日全天金币掉落翻倍，别忘了上线。",
        "image": "",
        "important": False,
        "min_code": 0,
        "max_code": 0,
        "lang": "zh",
        "enabled": False,
    },
]

FIELD_DOC = [
    ("id", "字符串", "公告唯一 ID，用于记录已读。不要重复，也不要修改已发布过的 ID"),
    ("title", "字符串", "标题"),
    ("date", "字符串", "日期，写成 2026-09-09 这种格式"),
    ("tag", "字符串", "标签：更新 / 活动 / 维护 / 补偿 / 公告"),
    ("content", "字符串", "正文。单元格内用 Alt+Enter 换行"),
    ("image", "字符串", "配图地址，可留空"),
    ("important", "TRUE/FALSE", "TRUE = 进游戏强制弹窗；FALSE = 只在公告列表里显示"),
    ("min_code", "数字", "最低版本号(versionCode)才显示，0 表示不限"),
    ("max_code", "数字", "最高版本号，0 表示不限"),
    ("lang", "字符串", "zh 中文 / en 英文"),
    ("enabled", "TRUE/FALSE", "FALSE = 这条不下发，等于临时下线"),
]


def _bool(v):
    if isinstance(v, bool):
        return v
    s = str(v).strip().lower()
    return s in ("true", "1", "yes", "y", "是", "真")


def _int(v, default=0):
    try:
        if v is None or str(v).strip() == "":
            return default
        return int(float(str(v)))
    except (TypeError, ValueError):
        return default


def _text(v):
    if v is None:
        return ""
    return str(v).replace("\r\n", "\n").replace("\r", "\n").strip()


def write_template():
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    if os.path.exists(XLSX):
        print("模板已存在，未覆盖：%s" % XLSX)
        return 0

    wb = Workbook()
    ws = wb.active
    ws.title = "公告"

    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="4F6BED")

    ws.append(FIELDS)
    for c in range(1, len(FIELDS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in SAMPLE_ROWS:
        ws.append([row.get(f, "") for f in FIELDS])

    widths = {"id": 16, "title": 28, "date": 12, "tag": 8, "content": 52,
              "image": 24, "important": 10, "min_code": 10, "max_code": 10,
              "lang": 8, "enabled": 10}
    for i, f in enumerate(FIELDS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(f, 14)
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=5).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 60
    ws.freeze_panes = "A2"

    doc = wb.create_sheet("字段说明")
    doc.append(["字段", "类型", "说明"])
    for c in range(1, 4):
        doc.cell(row=1, column=c).font = head_font
        doc.cell(row=1, column=c).fill = head_fill
    for row in FIELD_DOC:
        doc.append(list(row))
    doc.column_dimensions["A"].width = 14
    doc.column_dimensions["B"].width = 14
    doc.column_dimensions["C"].width = 70
    for r in range(2, doc.max_row + 1):
        doc.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(XLSX)
    print("已生成表格模板：%s" % XLSX)
    return 0


def read_rows():
    from openpyxl import load_workbook

    if not os.path.exists(XLSX):
        print("找不到表格：%s" % XLSX)
        print("先运行：python build_json.py --init")
        return None

    wb = load_workbook(XLSX, data_only=True)
    if "公告" not in wb.sheetnames:
        print('表格里缺少名为"公告"的工作表')
        return None
    ws = wb["公告"]

    header = [_text(c).lower() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    idx = {name: header.index(name) for name in FIELDS if name in header}
    missing = [f for f in FIELDS if f not in idx]
    if missing:
        print("表头缺少字段：%s" % ", ".join(missing))
        return None

    rows = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        if raw is None or all(v is None or str(v).strip() == "" for v in raw):
            continue
        row = {f: raw[idx[f]] for f in FIELDS}
        rows.append(row)
    return rows


def build(rows):
    items = []
    seen = set()
    for r in rows:
        if not _bool(r.get("enabled")):
            continue
        aid = _text(r.get("id"))
        if not aid:
            print("跳过一行：id 为空")
            continue
        if aid in seen:
            print("跳过重复 id：%s" % aid)
            continue
        seen.add(aid)
        items.append({
            "id": aid,
            "title": _text(r.get("title")),
            "date": _text(r.get("date")),
            "tag": _text(r.get("tag")),
            "content": _text(r.get("content")),
            "image": _text(r.get("image")),
            "important": _bool(r.get("important")),
            "min_code": _int(r.get("min_code"), 0),
            "max_code": _int(r.get("max_code"), 0),
            "lang": _text(r.get("lang")) or "zh",
        })

    items.sort(key=lambda x: x["date"], reverse=True)
    return {
        "schema": 1,
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "count": len(items),
        "items": items,
    }


def main():
    args = sys.argv[1:]
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        print("缺少 openpyxl，请先安装：pip install openpyxl")
        return 1

    if "--init" in args:
        return write_template()

    rows = read_rows()
    if rows is None:
        return 1

    data = build(rows)
    print("读到 %d 行，其中 %d 条启用" % (len(rows), data["count"]))

    if "--check" in args:
        print(json.dumps(data, ensure_ascii=False, indent=2))
        return 0

    with open(JSON_OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("已生成：%s" % JSON_OUT)
    for it in data["items"]:
        print("  [%s] %s (%s)%s" % (it["tag"], it["title"], it["date"], "  *强制*" if it["important"] else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
