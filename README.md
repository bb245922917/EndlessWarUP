# 无尽战争 - 游戏公告（在线热更新）

游戏启动时从这个仓库拉取 `announcement.json`，显示最新公告。改公告不用重新打包上架。

## 日常更新流程（三步）

1. 打开 `announcements.xlsx`，按行填（或改）公告，`enabled` 列填 `TRUE` 才下发
2. 双击 `生成公告.bat`，生成 `announcement.json`
3. 提交并推送：
   ```
   git add announcement.json announcements.xlsx
   git commit -m "更新公告"
   git push
   ```
   推送后约 1～2 分钟，游戏里就能看到新公告（游戏端有缓存，最迟下次启动生效）

## 游戏读取的地址

| 用途 | 地址 |
| --- | --- |
| 主地址（GitHub Pages，推荐） | `https://bb245922917.github.io/endlesswar-announcements/announcement.json` |
| 备用 1（jsDelivr CDN） | `https://cdn.jsdelivr.net/gh/bb245922917/endlesswar-announcements@main/announcement.json` |
| 备用 2（raw） | `https://raw.githubusercontent.com/bb245922917/endlesswar-announcements/main/announcement.json` |

主地址不通时游戏会自动依次尝试备用地址。

## 表格字段

完整说明在 `announcements.xlsx` 的「字段说明」工作表里，简要如下：

| 字段 | 说明 |
| --- | --- |
| id | 公告唯一 ID，用来记录"已读"，**不要改已经发过的** |
| title | 标题 |
| date | 日期 `2026-09-09` |
| tag | 更新 / 活动 / 维护 / 补偿 / 公告 |
| content | 正文，单元格内 Alt+Enter 换行 |
| image | 配图地址，可留空 |
| important | TRUE = 进游戏强制弹窗；FALSE = 只在公告列表里显示 |
| min_code | 最低版本号(versionCode)才显示，0 = 不限 |
| max_code | 最高版本号，0 = 不限 |
| lang | zh / en |
| enabled | FALSE = 这条不下发（临时下线用） |

## JSON 结构

```json
{
  "schema": 1,
  "updated": "2026-09-09 20:15:00",
  "count": 1,
  "items": [
    {
      "id": "A20260909001",
      "title": "新版本 0.1.6.4 更新内容",
      "date": "2026-09-09",
      "tag": "更新",
      "content": "1. 优化了低端机型的帧率表现\n2. 修复若干已知问题",
      "image": "",
      "important": true,
      "min_code": 185,
      "max_code": 0,
      "lang": "zh"
    }
  ]
}
```

## 注意事项

- 仓库是公开的，别放任何私密信息
- `announcement.json` 由脚本生成，不要手改（下次跑脚本会被覆盖）；要改就改表格
- 网络请求失败时游戏会静默跳过，不会影响正常进入游戏
