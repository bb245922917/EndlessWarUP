# 无尽战争 - 在线公告数据

游戏从这个仓库拉取 `announcement.json`，拿到**四国语言的公告文本**和**更新奖励的数量**。
改公告、改发多少金币钻石，都不用重新打包上架。

数据结构与游戏内结构体严格一致：
- `notice` 的四个字段 = `ST_公告结构体` 的 详情中文 / 详情英文 / 详情日文 / 详情韩文
- `rewards` 的字段 = 奖励结构体的 ID / 奖励名称 / 奖励数量（金币 ID=3，钻石 ID=4）

---

## 日常更新流程（三步）

1. 改三处（都是 UTF-8 文本，双击可用 Excel / 记事本打开）
   - `DT_公告数据表 - Sheet1.csv` — 四种语言的公告正文
   - `DT_更新奖励数据表 - Sheet1.csv` — 金币、钻石的数量
   - `version_code.txt` — 这次的 versionCode，纯数字（当前 `185`）
2. 双击 `生成公告.bat`，生成 `announcement.json`
3. 提交推送：
   ```
   git add announcement.json "DT_公告数据表 - Sheet1.csv" "DT_更新奖励数据表 - Sheet1.csv" version_code.txt
   git commit -m "更新公告"
   git push
   ```
   推送后：raw 地址几乎立即生效；主地址 jsDelivr 有最长约 12 小时缓存（可到
   `https://purge.jsdelivr.net` 手动清缓存）。游戏端每次启动都会拉，最迟下次启动拿得到。

---

## 游戏读取的地址

仓库：<https://github.com/bb245922917/EndlessWarUP>（公开，默认分支 `master`）

| 用途 | 地址 | 状态 |
| --- | --- | --- |
| **主地址（jsDelivr CDN）** | `https://cdn.jsdelivr.net/gh/bb245922917/EndlessWarUP@master/announcement.json` | 已实测 200 |
| 备用 1（raw） | `https://raw.githubusercontent.com/bb245922917/EndlessWarUP/master/announcement.json` | 已实测 200 |
| 备用 2（GitHub Pages） | `https://bb245922917.github.io/EndlessWarUP/announcement.json` | 需先开 Pages，当前 404 |

**蓝图里现在填的是主地址（jsDelivr）**，无需任何额外设置，改完表格推送即可生效。

想让更新更及时（jsDelivr 有约 12 小时缓存）：
仓库 `Settings` → `Pages` → `Deploy from a branch` → 分支 `master` → `/ (root)` → `Save`，
开启后把蓝图里的 URL 换成上表的「备用 2」即可。

---

## JSON 结构

```json
{
  "schema": 1,
  "version_name": "V 0.1.6.3",
  "version_code": 185,
  "updated": "2026-09-10 06:10:42",
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

## CSV 说明

**`DT_公告数据表 - Sheet1.csv`**（只保留一行数据）

| 列 | 说明 |
| --- | --- |
| 编号 | 行号，对应 JSON 里的 `notice.Name`，一般填 `1` |
| 详情英文 / 详情中文 / 详情日文 / 详情韩文 | 四种语言正文，列名与 `ST_公告结构体` 字段同名 |
| 版本号 | **不用单独列** —— 脚本从「详情中文」第一行 `版本号:V x.y.z` 自动提取 |

单元格内换行直接回车即可，CSV 里就是一个 `\n`。

**`DT_更新奖励数据表 - Sheet1.csv`**（每种奖励一行）

| 列 | 说明 |
| --- | --- |
| --- | 行号，占位用 |
| 奖励名称 | 金币 / 钻石 |
| ID | 金币 = 3，钻石 = 4，别填错，蓝图按这个分流 |
| 奖励数量 | 这次发多少，改这里即时生效，不用重新打包 |

**`version_code.txt`**：只有一个数字，这次的 versionCode（当前 `185`）。
游戏用它判断是不是新版本，发新版时记得改。

---

## 注意

- 这个仓库是公开的，别放任何私密信息
- **⚠️ CSV 必须存成 UTF-8**。Excel 另存时要选「CSV UTF-8(逗号分隔)」。
  存成默认的 GBK/ANSI 的话，**韩文会立刻变成 `???` 且无法恢复**（谚文不在 GBK 字符集里）。
  中文和日文假名在 GBK 里能存，所以只有韩文会坏，很容易漏掉。
  脚本检测到韩文里出现 `?` 会在生成时打印警告。
- `announcement.json` 由脚本生成，不要手改（下次跑脚本会覆盖）；要改就改 CSV
- 网络拿不到数据时，游戏会回退用本地的 `DT_公告数据表` / `DT_更新奖励数据表`，不会白屏也不会卡住
