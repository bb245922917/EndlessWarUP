#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公告 CSV -> announcement.json

用法:
    python build_json.py            读取 CSV，生成 announcement.json
    python build_json.py --check    只打印将要生成的 JSON，不写文件
    python build_json.py --fix      修复 CSV 编码（GBK->UTF-8）并恢复被破坏的韩文
    python build_json.py --init     重建两份 CSV 模板

数据源（都在本目录）:
    DT_公告数据表 - Sheet1.csv       四国语言公告
                                    列：编号,详情英文,详情中文,详情日文,详情韩文
    DT_更新奖励数据表 - Sheet1.csv   这次更新发的奖励
                                    列按 ST_日常任务奖励结构体：
                                    ---,奖励名称,ID,奖励数量,物品图标,奖励?,
                                    说明,中文说明,已领取,活跃值,日文说明,韩文说明
    version_code.txt                这次的 versionCode（纯数字，如 185）

编码问题（重要）:
    CSV 必须是 UTF-8。Excel 默认另存是 GBK/ANSI，而**谚文不在 GBK 字符集**，
    一存就变成 ??? 且无法恢复（中文、日文假名在 GBK 里能存，所以只有韩文会坏）。
    本脚本会自动识别 GBK 并告警；用 --fix 可转成 UTF-8 并恢复韩文。
    读写统一按 utf-8-sig（带 BOM），Excel 双击打开不乱码。
"""

import csv
import io
import json
import os
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
NOTICE_CSV = os.path.join(HERE, "DT_公告数据表 - Sheet1.csv")
REWARD_CSV = os.path.join(HERE, "DT_更新奖励数据表 - Sheet1.csv")
CODE_TXT = os.path.join(HERE, "version_code.txt")
JSON_OUT = os.path.join(HERE, "announcement.json")

# 与 ST_公告结构体 字段一一对应（顺序按工程内 DataTable 的字段顺序）
NOTICE_LANGS = ["详情英文", "详情中文", "详情日文", "详情韩文"]
# 输出 JSON 时的顺序（中文在前方便阅读；蓝图按字段名取值，顺序无影响）
OUT_ORDER = ["详情中文", "详情英文", "详情日文", "详情韩文"]

# 韩文模板（与工程内 DT_公告数据表 的原文一致，版本号动态替换）
KO_TEMPLATE = (
    "버전:{ver}\n"
    "1.스킬 피해 범위 강화；\n"
    "2.자원 획득 용이；\n"
    "3.광고 응답 속도 향상；\n"
    "4.여러 버그 수정。"
)

SAMPLE_NOTICE = {
    "编号": "1",
    "详情英文": "Version:V 0.1.6.3\n1.Enhance skill damage range;\n2.Make resources easier to obtain;\n3.Improve ad response speed;\n4.Fix several bugs.",
    "详情中文": "版本号:V 0.1.6.3\n1.增强技能伤害范围；\n2.资源获得更容易；\n3.提升广告响应速度：\n4.修复若干Bug。",
    "详情日文": "バージョン:V 0.1.6.3\n1.スキルのダメージ範囲を強化；\n2.リソースの獲得を容易に；\n3.広告の応答速度を向上；\n4.いくつかのバグを修正。",
    "详情韩文": KO_TEMPLATE.format(ver="V 0.1.6.3"),
}

# 奖励表列：与 ST_日常任务奖励结构体 完全一致
REWARD_FIELDS = [
    "---", "奖励名称", "ID", "奖励数量", "物品图标", "奖励?",
    "说明", "中文说明", "已领取", "活跃值", "日文说明", "韩文说明",
]

_NS = '[8F3A4F7F78E52C175B3305EF4AA5E433]'
_ICON = "/Script/Engine.Texture2D'/Game/BreakTheBricks/Art/UI/主页/Texture/"


def _loc(key, val):
    return 'NSLOCTEXT("' + _NS + '", "' + key + '", "' + val + '")'


SAMPLE_REWARDS = [
    {
        "---": "1",
        "奖励名称": "金币",
        "ID": "3",
        "奖励数量": "1000.000000",
        "物品图标": _ICON + "T_UI_Item_Gtid_金币_01a.T_UI_Item_Gtid_金币_01a'",
        "奖励?": "False",
        "说明": _loc("E35E5A4345B6FDC8C025618594E31B76", "Gold"),
        "中文说明": _loc("DA80D0DD429DDC149D37B591EBEEEBB8", "金币"),
        "已领取": "False",
        "活跃值": "0.000000",
        "日文说明": _loc("E089AACF4A1D47D04315DFBB39E3AF4B", "ゴールド"),
        "韩文说明": _loc("2D0324114E3A67D5C24350A2CA930399", "골드"),
    },
    {
        "---": "2",
        "奖励名称": "钻石",
        "ID": "4",
        "奖励数量": "600.000000",
        "物品图标": _ICON + "T_UI_Item_Gtid_钻石_01b.T_UI_Item_Gtid_钻石_01b'",
        "奖励?": "False",
        "说明": _loc("1AD715994D385EA2CB6D94B3D92DFF8A", "Diamonds"),
        "中文说明": _loc("E5C20F38442990207213DFA25E9324C7", "钻石"),
        "已领取": "False",
        "活跃值": "0.000000",
        "日文说明": _loc("7B65EAA447EB125C669D5DBCB3AB1FDA", "ダイヤ"),
        "韩文说明": _loc("1F28F7774025D33FDEC4AA8652809EB6", "다이아"),
    },
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


def read_csv_any(path):
    """自动识别编码读 CSV，返回 (字段名, 行列表, 编码)"""
    if not os.path.exists(path):
        return None, None, None
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "gbk", "cp936"):
        try:
            text = raw.decode(enc)
            rows = list(csv.DictReader(io.StringIO(text)))
            fields = list(rows[0].keys()) if rows else []
            return fields, rows, enc
        except UnicodeDecodeError:
            continue
    raise SystemExit("无法识别编码: %s" % path)


def write_csv_utf8(path, fields, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow(r)


def pick_version_name(cn_text):
    """从中文正文第一行「版本号:V x.y.z」里取版本名"""
    for line in cn_text.splitlines():
        line = line.strip()
        for prefix in ("版本号:", "版本号："):
            if line.startswith(prefix):
                return line[len(prefix):].strip()
    return ""


def fix_notice_csv():
    """把公告 CSV 转成 UTF-8，并恢复被 GBK 破坏的韩文"""
    fields, rows, enc = read_csv_any(NOTICE_CSV)
    if rows is None:
        print("找不到:", NOTICE_CSV)
        return 1
    print("当前编码:", enc)

    changed_enc = enc not in ("utf-8-sig",)
    fixed_ko = 0
    for r in rows:
        cn = _text(r.get("详情中文"))
        ko = _text(r.get("详情韩文"))
        if "?" in ko:
            r["详情韩文"] = KO_TEMPLATE.format(ver=pick_version_name(cn))
            fixed_ko += 1
        for k in NOTICE_LANGS:
            if r.get(k):
                r[k] = _text(r.get(k))

    write_csv_utf8(NOTICE_CSV, fields, rows)
    print("已按 UTF-8(with BOM) 重写:", NOTICE_CSV)
    if changed_enc:
        print("  编码已从 %s 转换为 UTF-8" % enc)
    if fixed_ko:
        print("  已恢复 %d 行被破坏的韩文" % fixed_ko)
    else:
        print("  韩文无需恢复")
    return 0


def fix_reward_csv():
    """把奖励 CSV 转 UTF-8，并恢复被 GBK 破坏的韩文/日文说明"""
    fields, rows, enc = read_csv_any(REWARD_CSV)
    if rows is None:
        print("找不到:", REWARD_CSV)
        return 0
    print("奖励表当前编码:", enc)

    sample_by_id = {s["ID"]: s for s in SAMPLE_REWARDS}
    fixed = 0
    for r in rows:
        s = sample_by_id.get(_text(r.get("ID")))
        if not s:
            continue
        for k in ("韩文说明", "日文说明"):
            v = _text(r.get(k))
            if v and "?" in v:
                r[k] = s[k]
                fixed += 1

    write_csv_utf8(REWARD_CSV, fields, rows)
    print("  已按 UTF-8(with BOM) 重写:", REWARD_CSV)
    print("  恢复字段数:", fixed)
    return 0


def write_template():
    if not os.path.exists(NOTICE_CSV):
        write_csv_utf8(NOTICE_CSV, ["编号"] + NOTICE_LANGS, [SAMPLE_NOTICE])
        print("已生成:", NOTICE_CSV)
    else:
        print("已存在，跳过:", NOTICE_CSV)

    if not os.path.exists(REWARD_CSV):
        write_csv_utf8(REWARD_CSV, REWARD_FIELDS, SAMPLE_REWARDS)
        print("已生成:", REWARD_CSV)
    else:
        print("已存在，跳过:", REWARD_CSV)

    if not os.path.exists(CODE_TXT):
        with open(CODE_TXT, "w", encoding="utf-8", newline="") as f:
            f.write("185")
        print("已生成:", CODE_TXT)
    return 0


def read_version_code():
    if os.path.exists(CODE_TXT):
        with open(CODE_TXT, "r", encoding="utf-8-sig") as f:
            return _int(_text(f.read()), 0)
    if os.path.exists(JSON_OUT):
        try:
            with open(JSON_OUT, "r", encoding="utf-8") as f:
                return _int(json.load(f).get("version_code"), 0)
        except Exception:
            return 0
    return 0


def build():
    n_fields, n_rows, n_enc = read_csv_any(NOTICE_CSV)
    if n_rows is None:
        print("找不到公告表：%s" % NOTICE_CSV)
        print("先运行：python build_json.py --init")
        return None
    if not n_rows:
        print("公告表是空的：%s" % NOTICE_CSV)
        return None

    missing = [c for c in NOTICE_LANGS if c not in (n_fields or [])]
    if missing:
        print("公告表缺少列：%s" % ", ".join(missing))
        print("当前列：%s" % ", ".join(n_fields or []))
        return None

    warns = []
    if n_enc not in ("utf-8-sig", "utf-8"):
        warns.append("公告表是 %s 编码，不是 UTF-8，韩文可能已损坏。运行 python build_json.py --fix" % n_enc)

    n = n_rows[0]
    langs = {k: _text(n.get(k)) for k in NOTICE_LANGS}
    if "?" in langs["详情韩文"]:
        # CSV 被存成 GBK，谚文已变成 ?。先用内置正确韩文兜底，保证线上不会发出 ??
        langs["详情韩文"] = KO_TEMPLATE.format(ver=pick_version_name(langs["详情中文"]))
        warns.append(
            "详情韩文 里出现 ?（存成了 GBK，谚文已丢失）。本次已用内置正确韩文写入 JSON，"
            "但 CSV 本身还是坏的，请关闭 Excel 后运行 python build_json.py --fix"
        )

    notice = {"Name": _text(n.get("编号")) or "1"}
    for k in OUT_ORDER:
        notice[k] = langs[k]

    rewards = []
    r_fields, r_rows, r_enc = read_csv_any(REWARD_CSV)
    if r_rows:
        for r in r_rows:
            if not _text(r.get("ID")):
                continue
            rewards.append({
                "ID": _text(r.get("ID")),
                "奖励名称": _text(r.get("奖励名称")),
                "奖励数量": _int(r.get("奖励数量"), 0),
            })
    if r_rows and r_enc not in ("utf-8-sig", "utf-8"):
        warns.append("奖励表是 %s 编码，建议转 UTF-8" % r_enc)

    data = {
        "schema": 1,
        "version_name": pick_version_name(langs["详情中文"]),
        "version_code": read_version_code(),
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "notice": notice,
        "rewards": rewards,
    }
    if warns:
        data["_warning"] = warns
    return data


def main():
    args = sys.argv[1:]

    if "--init" in args:
        return write_template()
    if "--fix" in args:
        fix_notice_csv()
        return fix_reward_csv()

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
    for k in OUT_ORDER:
        print("  %s：%d 字" % (k, len(data["notice"][k])))
    for r in data["rewards"]:
        print("  奖励：%s x %d (ID=%s)" % (r["奖励名称"], r["奖励数量"], r["ID"]))
    for w in data.get("_warning", []):
        print("  [警告] %s" % w)
    return 0


if __name__ == "__main__":
    sys.exit(main())
