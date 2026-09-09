#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公告 CSV -> announcement.json

用法:
    python build_json.py            读取 CSV，生成 announcement.json
    python build_json.py --check    只打印将要生成的 JSON，不写文件
    python build_json.py --init     重建两份 CSV 模板（会覆盖已有模板）

数据源（都在本目录）:
    DT_公告数据表 - Sheet1.csv       四国语言公告
                                    列：编号,详情英文,详情中文,详情日文,详情韩文
    DT_更新奖励数据表 - Sheet1.csv   这次更新发的奖励
                                    列：---,奖励名称,ID,奖励数量
    version_code.txt                这次的 versionCode（纯数字，如 185）

注意：
    CSV 必须是 UTF-8 编码。用 Excel 另存时请选「CSV UTF-8(逗号分隔)」，
    否则韩文会变成 ???（谚文不在 GBK 字符集里，保存即丢失，无法恢复）。
    脚本读写都按 utf-8-sig（带 BOM）处理，Excel 双击打开不会乱码。

生成的 JSON 里 notice 的四个字段名与 ST_公告结构体 完全一致，
蓝图里可直接 Make ST_公告结构体，现有逻辑不用改。
"""

import csv
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

SAMPLE_NOTICE = {
    "编号": "1",
    "详情英文": "Version:V 0.1.6.3\n1.Enhance skill damage range;\n2.Make resources easier to obtain;\n3.Improve ad response speed;\n4.Fix several bugs.",
    "详情中文": "版本号:V 0.1.6.3\n1.增强技能伤害范围；\n2.资源获得更容易；\n3.提升广告响应速度：\n4.修复若干Bug。",
    "详情日文": "バージョン:V 0.1.6.3\n1.スキルのダメージ範囲を強化；\n2.リソースの獲得を容易に；\n3.広告の応答速度を向上；\n4.いくつかのバグを修正。",
    "详情韩文": "버전:V 0.1.6.3\n1.스킬 피해 범위 강화；\n2.자원 획득 용이；\n3.광고 응답 속도 향상；\n4.여러 버그 수정。",
}

SAMPLE_REWARDS = [
    {"---": "1", "奖励名称": "金币", "ID": "3", "奖励数量": "1000"},
    {"---": "2", "奖励名称": "钻石", "ID": "4", "奖励数量": "600"},
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


def read_csv(path):
    """按 utf-8-sig 读 CSV，返回 (字段名列表, 行字典列表)"""
    if not os.path.exists(path):
        return None, None
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    if not rows:
        return [], []
    return list(rows[0].keys()), rows


def write_template():
    """重建两份 CSV 模板 + version_code.txt"""
    with open(NOTICE_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["编号"] + NOTICE_LANGS)
        w.writeheader()
        w.writerow(SAMPLE_NOTICE)
    print("已生成:", NOTICE_CSV)

    with open(REWARD_CSV, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["---", "奖励名称", "ID", "奖励数量"])
        w.writeheader()
        for r in SAMPLE_REWARDS:
            w.writerow(r)
    print("已生成:", REWARD_CSV)

    if not os.path.exists(CODE_TXT):
        with open(CODE_TXT, "w", encoding="utf-8", newline="") as f:
            f.write("185")
        print("已生成:", CODE_TXT)
    return 0


def pick_version_name(cn_text):
    """从中文正文第一行「版本号:V x.y.z」里取版本名"""
    for line in cn_text.splitlines():
        line = line.strip()
        for prefix in ("版本号:", "版本号："):
            if line.startswith(prefix):
                return line[len(prefix):].strip()
    return ""


def read_version_code():
    if os.path.exists(CODE_TXT):
        with open(CODE_TXT, "r", encoding="utf-8-sig") as f:
            return _int(_text(f.read()), 0)
    # 没有配置文件就沿用上一版 JSON 里的值
    if os.path.exists(JSON_OUT):
        try:
            with open(JSON_OUT, "r", encoding="utf-8") as f:
                return _int(json.load(f).get("version_code"), 0)
        except Exception:
            return 0
    return 0


def build():
    n_fields, n_rows = read_csv(NOTICE_CSV)
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

    n = n_rows[0]
    langs = {k: _text(n.get(k)) for k in NOTICE_LANGS}

    warned = None
    if "?" in langs["详情韩文"]:
        warned = "详情韩文 里出现 ?，可能是用非 UTF-8 保存导致谚文丢失，请检查"

    notice = {"Name": _text(n.get("编号")) or "1"}
    for k in OUT_ORDER:
        notice[k] = langs[k]

    rewards = []
    r_fields, r_rows = read_csv(REWARD_CSV)
    if r_rows:
        for r in r_rows:
            if not _text(r.get("ID")):
                continue
            rewards.append({
                "ID": _text(r.get("ID")),
                "奖励名称": _text(r.get("奖励名称")),
                "奖励数量": _int(r.get("奖励数量"), 0),
            })

    data = {
        "schema": 1,
        "version_name": pick_version_name(langs["详情中文"]),
        "version_code": read_version_code(),
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "notice": notice,
        "rewards": rewards,
    }
    if warned:
        data["_warning"] = warned
    return data


def main():
    args = sys.argv[1:]

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
    for k in OUT_ORDER:
        print("  %s：%d 字" % (k, len(data["notice"][k])))
    for r in data["rewards"]:
        print("  奖励：%s x %d (ID=%s)" % (r["奖励名称"], r["奖励数量"], r["ID"]))
    if data.get("_warning"):
        print("  [警告] %s" % data["_warning"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
