#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公告表格 -> announcement.json

用法:
    python build_json.py            读取 announcements.xlsx，生成 announcement.json
    python build_json.py --init     重建 announcements.xlsx 模板（覆盖已有模板）
    python build_json.py --check    只打印将要生成的 JSON，不写文件

表格有两个工作表:
    [公告]  一行，四个语言字段，字段名与 ST_公告结构体 保持一致
    [奖励]  每行一种奖励，字段与 ST_日常任务奖励结构体 的对应字段保持一致

生成的 JSON 里 notice 部分字段名与 ST_公告结构体 完全相同，
蓝图里可直接 Make ST_公告结构体 填进去，现有逻辑不用改。
"""

import json
import os
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "announcements.xlsx")
JSON_OUT = os.path.join(HERE, "announcement.json")

# 与 ST_公告结构体 字段一一对应
NOTICE_LANGS = ["详情中文", "详情英文", "详情日文", "详情韩文"]
NOTICE_COLS = ["版本名", "版本号"] + NOTICE_LANGS

REWARD_COLS = ["ID", "奖励名称", "奖励数量", "备注"]

SAMPLE_NOTICE = {
    "版本名": "V 0.1.6.3",
    "版本号": 185,
    "详情中文": "版本号:V 0.1.6.3\r\n1.增强技能伤害范围；\r\n2.资源获得更容易；\r\n3.提升广告响应速度：\r\n4.修复若干Bug。",
    "详情英文": "Version:V 0.1.6.3\r\n1.Enhance skill damage range;\r\n2.Make resources easier to obtain;\r\n3.Improve ad response speed;\r\n4.Fix several bugs.",
    "详情日文": "バージョン:V 0.1.6.3\r\n1.スキルのダメージ範囲を強化；\r\n2.リソースの獲得を容易に；\r\n3.広告の応答速度を向上；\r\n4.いくつかのバグを修正。",
    "详情韩文": "버전:V 0.1.6.3\r\n1.스킬 피해 범위 강화；\r\n2.자원 획득 용이；\r\n3.광고 응답 속도 향상；\r\n4.여러 버그 수정。",
}

SAMPLE_REWARDS = [
    {"ID": "3", "奖励名称": "金币", "奖励数量": 1000, "备注": "更新奖励"},
    {"ID": "4", "奖励名称": "钻石", "奖励数量": 600, "备注": "更新奖励"},
]

DOC = [
    ("公告", "版本名", "字符串", "显示用的版本号文字，例如 V 0.1.6.3"),
    ("公告", "版本号", "数字", "versionCode，用来判断是不是新版本（185 = 0.1.6.3）"),
    ("公告", "详情中文", "字符串", "对应 ST_公告结构体.详情中文，单元格内 Alt+Enter 换行"),
    ("公告", "详情英文", "字符串", "对应 ST_公告结构体.详情英文"),
    ("公告", "详情日文", "字符串", "对应 ST_公告结构体.详情日文"),
    ("公告", "详情韩文", "字符串", "对应 ST_公告结构体.详情韩文"),
    ("奖励", "ID", "字符串", "对应奖励结构体的 ID，金币=3，钻石=4，不要用错"),
    ("奖励", "奖励名称", "字符串", "金币 / 钻石"),
    ("奖励", "奖励数量", "数字", "这次更新发放的数量，改这里即可，不用改游戏"),
    ("奖励", "备注", "字符串", "给自己看的说明，不会进 JSON"),
]


def _text(v):
    if v is None:
        return ""
    return str(v).replace("\r\n", "\n").replace("\r", "\n").strip()


def _int(v, default=0):
    try:
        if v is None or str(v).strip() == "":
            return default
        return int(float(str(v)))
    except (TypeError, ValueError):
        return default


def write_template():
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    if os.path.exists(XLSX):
        os.remove(XLSX)

    wb = Workbook()
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="4F6BED")

    ws = wb.active
    ws.title = "公告"
    ws.append(NOTICE_COLS)
    ws.append([SAMPLE_NOTICE.get(c, "") for c in NOTICE_COLS])
    for c in range(1, len(NOTICE_COLS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    widths = {"版本名": 16, "版本号": 10, "详情中文": 58, "详情英文": 58, "详情日文": 58, "详情韩文": 58}
    for i, f in enumerate(NOTICE_COLS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(f, 16)
        if f in NOTICE_LANGS:
            ws.cell(row=2, column=i).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 110
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("奖励")
    ws2.append(REWARD_COLS)
    for row in SAMPLE_REWARDS:
        ws2.append([row.get(c, "") for c in REWARD_COLS])
    for c in range(1, len(REWARD_COLS) + 1):
        cell = ws2.cell(row=1, column=c)
        cell.font = head_font
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for i, w in enumerate([10, 14, 12, 24], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"

    ws3 = wb.create_sheet("字段说明")
    ws3.append(["工作表", "字段", "类型", "说明"])
    for c in range(1, 5):
        ws3.cell(row=1, column=c).font = head_font
        ws3.cell(row=1, column=c).fill = head_fill
    for row in DOC:
        ws3.append(list(row))
    for col, w in zip("ABCD", [10, 14, 10, 70]):
        ws3.column_dimensions[col].width = w
    for r in range(2, ws3.max_row + 1):
        ws3.cell(row=r, column=4).alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(XLSX)
    print("已生成表格模板：%s" % XLSX)
    return 0


def _read_sheet(wb, sheet, cols):
    if sheet not in wb.sheetnames:
        print("缺少工作表：%s" % sheet)
        return None
    ws = wb[sheet]
    header = [_text(c) for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    idx = {}
    for name in cols:
        if name in header:
            idx[name] = header.index(name)
    missing = [c for c in cols if c not in idx]
    if missing:
        print("%s 表头缺少字段：%s" % (sheet, ", ".join(missing)))
        return None
    rows = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        if raw is None or all(v is None or str(v).strip() == "" for v in raw):
            continue
        rows.append({c: raw[idx[c]] for c in cols})
    return rows


def build():
    from openpyxl import load_workbook

    if not os.path.exists(XLSX):
        print("找不到表格：%s" % XLSX)
        print("先运行：python build_json.py --init")
        return None

    wb = load_workbook(XLSX, data_only=True)
    notice_rows = _read_sheet(wb, "公告", NOTICE_COLS)
    reward_rows = _read_sheet(wb, "奖励", REWARD_COLS)
    if not notice_rows or not reward_rows:
        return None

    n = notice_rows[0]
    notice = {"Name": "1"}
    for lang in NOTICE_LANGS:
        notice[lang] = _text(n.get(lang))

    rewards = []
    for r in reward_rows:
        rewards.append({
            "ID": _text(r.get("ID")),
            "奖励名称": _text(r.get("奖励名称")),
            "奖励数量": _int(r.get("奖励数量"), 0),
        })

    return {
        "schema": 1,
        "version_name": _text(n.get("版本名")),
        "version_code": _int(n.get("版本号"), 0),
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "notice": notice,
        "rewards": rewards,
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

    data = build()
    if data is None:
        return 1

    if "--check" in args:
        print(json.dumps(data, ensure_ascii=False, indent=2))
        return 0

    with open(JSON_OUT, "w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("已生成：%s" % JSON_OUT)
    print("  版本：%s (versionCode %s)" % (data["version_name"], data["version_code"]))
    for k in NOTICE_LANGS:
        print("  %s：%d 字" % (k, len(data["notice"][k])))
    for r in data["rewards"]:
        print("  奖励：%s x %d (ID=%s)" % (r["奖励名称"], r["奖励数量"], r["ID"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
