# 无尽战争 - 在线公告数据

游戏从这个仓库拉取 `announcement.json`，拿到**四国语言的公告文本**和**更新奖励的数量**。
改公告、改发多少金币钻石，都不用重新打包上架。

数据结构与游戏内结构体严格一致：
- `notice` 的四个字段 = `ST_公告结构体` 的 详情中文 / 详情英文 / 详情日文 / 详情韩文
- `rewards` 的字段 = 奖励结构体的 ID / 奖励名称 / 奖励数量（金币 ID=3，钻石 ID=4）

---

## 数据源（XLSX，不再用 CSV）

> 为什么是 XLSX：CSV 在 Excel 里默认存成 GBK/ANSI，而**谚文不在 GBK 字符集**，一存就变 `?` 且无法恢复。
> XLSX 是 Unicode 原生存储（Office Open XML），中/日/韩文都不会丢，彻底根治乱码。
> 生成脚本需要 `openpyxl`（本机 venv 已装；如需重装：`pip install openpyxl`）。

- `DT_公告数据表.xlsx` — 四种语言的公告正文
- `DT_更新奖励数据表.xlsx` — 金币、钻石的数量
- `version_code.txt` — 这次的 versionCode，纯数字（当前 `185`）

---

## 日常更新流程（三步）

1. 改两处（用 Excel 打开下面的 XLSX，直接编辑文字，**另存为时选 .xlsx，不要选 CSV**）
   - `DT_公告数据表.xlsx` — 四种语言的公告正文
   - `DT_更新奖励数据表.xlsx` — 金币、钻石的数量
   - `version_code.txt` — 这次的 versionCode，纯数字（当前 `185`）
2. 双击 `生成公告.bat`，生成 `announcement.json`
3. 提交推送：
   ```
   git add announcement.json "DT_公告数据表.xlsx" "DT_更新奖励数据表.xlsx" version_code.txt
   git commit -m "更新公告"
   git push
   ```
   推送后 GitHub Pages 约 1~2 分钟生效，游戏端下次启动即拉到最新（无 12 小时缓存）。

---

## 游戏读取的地址

仓库：<https://github.com/bb245922917/EndlessWarUpdate>（公开，默认分支 `master`）

| 用途 | 地址 | 状态 |
| --- | --- | --- |
| **主地址（GitHub Pages）** | `https://bb245922917.github.io/EndlessWarUpdate/announcement.json` | 已开 Pages，推送后 1~2 分钟生效 |
| 备用 1（raw） | `https://raw.githubusercontent.com/bb245922917/EndlessWarUpdate/master/announcement.json` | 已实测 200 |
| 备用 2（jsDelivr CDN） | `https://cdn.jsdelivr.net/gh/bb245922917/EndlessWarUpdate@master/announcement.json` | 已实测 200，但有最长约 12 小时分支缓存 |

**蓝图里现在填的是 GitHub Pages 主地址**。改完表格推送即可生效，不用重打包。

---

## JSON 结构

```json
{
  "schema": 1,
  "version_name": "V 0.1.6.3",
  "version_code": 185,
  "updated": "2026-09-11 02:35:31",
  "notice": {
    "Name": "1",
    "详情中文": "版本号:V 0.1.6.3\n1.增强技能伤害范围；\n2.资源获得更容易；\n3.提升广告响应速度：\n4.修复若干Bug。",
    "详情英文": "Version:V 0.1.6.3\n1.Enhance skill damage range;\n...",
    "详情日文": "...",
    "详情韩文": "..."
  },
  "rewards": [
    { "ID": "3", "奖励名称": "金币", "奖励数量": 1000 },
    { "ID": "4", "奖励名称": "钻石", "奖励数量": 600 }
  ]
}
```

---

## XLSX 说明

**`DT_公告数据表.xlsx`**（只保留一行数据）

| 列 | 说明 |
| --- | --- |
| 编号 | 行号，对应 JSON 里的 `notice.Name`，一般填 `1` |
| 详情英文 / 详情中文 / 详情日文 / 详情韩文 | 四种语言正文，列名与 `ST_公告结构体` 字段同名 |
| 版本号 | **不用单独列** —— 脚本从「详情中文」第一行 `版本号:V x.y.z` 自动提取 |

单元格内换行直接回车即可，XLSX 里就是 `\n`。

**`DT_更新奖励数据表.xlsx`**（每种奖励一行）

列名与工程内 **`ST_日常任务奖励结构体`** 完全一致，就是 UE 直接导出的格式，可以原样导回 `DT_更新奖励数据表`：

| 列 | 类型 | 说明 |
| --- | --- | --- |
| --- | 行名 | UE 导出的行号占位列，填 `1` / `2` |
| 奖励名称 | String | 金币 / 钻石 |
| ID | String | **金币 = 3，钻石 = 4**，蓝图按这个分流，别填错 |
| 奖励数量 | int | 这次发多少，改这里即时生效，不用重新打包 |
| 物品图标 | Texture2D | 图标资源路径，一般不用改 |
| 奖励? | bool | 是否计入奖励 |
| 说明 | FText | 英文说明（NSLOCTEXT） |
| 中文说明 | FText | 中文说明（NSLOCTEXT） |
| 已领取 | bool | 领取状态 |
| 活跃值 | int | 活跃度 |
| 日文说明 | FText | 日文说明（NSLOCTEXT） |
| 韩文说明 | FText | 韩文说明（NSLOCTEXT） |

真正进 JSON 的只有 `ID` / `奖励名称` / `奖励数量` 三列，其余列保持与 UE 一致，方便导回。

**`version_code.txt`**：只有一个数字，这次的 versionCode（当前 `185`）。
游戏用它判断是不是新版本，发新版时记得改。

---

## 注意

- 这个仓库是公开的，别放任何私密信息
- **韩文乱码**：已通过改用 XLSX 彻底解决（XLSX 是 Unicode 存储，Excel 编辑不会丢谚文）。
  只有一种例外——如果某个 XLSX 曾被当成 CSV 另存过，韩文才会变 `?`；
  这时先关掉 Excel / WPS，双击 `修复编码.bat`（等价于 `python build_json.py --fix`）即可用内置正确韩文兜底。
- `announcement.json` 由脚本生成，不要手改（下次跑脚本会覆盖）；要改就改 XLSX
- 网络拿不到数据时，游戏会回退用本地的 `DT_公告数据表` / `DT_更新奖励数据表`，不会白屏也不会卡住
