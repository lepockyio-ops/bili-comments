---
name: bili-comments
description: 采集 B 站视频的所有评论并导出为 Google Sheets 友好的 XLSX（含中→日自动翻译公式）。当用户提供 B 站视频链接或 bvid，说"抓评论""收集评论""导出到表格""翻译评论""看看别人怎么评的"等意图时使用。为日本创作者/曲师看粉丝反馈场景设计。
---

# BiliComments — B 站评论收集器

将 B 站视频的评论全量抓下来 → 导出成 XLSX → 上传 Google Sheets 后自动出现日语翻译。

## 何时使用

- 用户丢来一个 B 站视频链接或 bvid，问"评论怎么样""能看看评论吗""导出评论"
- 用户想给日本曲师/绘师/合作方看粉丝反馈
- 用户要做爆款曲的评论区分析、切片辅助
- 用户想快速搞清楚"负面评论/求授权/求歌词"这类特定内容有多少

## 优先执行方式

```powershell
Set-Location "C:\Users\何\Documents\Claude\Projects\MCN创业\bili-comments"
.\run.ps1 <bvid_或_url> [选项]
```

`run.ps1` 自动处理 UTF-8 + 依赖安装。

## 常用命令

```powershell
# 抓热度前 N 页主评论（默认 hot 排序，抓完所有页）
.\run.ps1 BV1FDNJ6BE4j

# 只抓前 5 页（快速尝鲜，100 条）
.\run.ps1 BV1FDNJ6BE4j --max-pages 5

# 按时间新→旧
.\run.ps1 BV1FDNJ6BE4j --sort new

# 含所有楼中楼（子评论），完整但慢
.\run.ps1 BV1FDNJ6BE4j --include-replies

# 指定输出路径
.\run.ps1 BV1FDNJ6BE4j --out C:\Users\me\Desktop\comments.xlsx

# 翻成英语/韩语
.\run.ps1 BV1FDNJ6BE4j --target-lang en
.\run.ps1 BV1FDNJ6BE4j --target-lang ko
```

## 输出文件

默认保存到 `data/<bvid>.xlsx`，含 3 个 Sheet：

1. **评论**：所有评论（15 列，含 GOOGLETRANSLATE 翻译公式）
2. **视频信息**：视频元数据 + 采集统计
3. **Top 20 高赞**：最热的 20 条主评论，便于扫读

## 关键行为

- **翻译列是 Google Sheets 公式**：`=GOOGLETRANSLATE(H2,"zh-CN","ja")`
  - **上传到 Google Sheets 后自动翻译**
  - 用本地 Excel/WPS 打开会显示 `#NAME?` 错误（属正常，因为是 Google 独有函数）
- **UP 主评论行高亮**：黄色背景，一眼识别曲师本人的置顶/回复
- **风控友好**：每页请求间随机 sleep 1.5-3s
- **首页含置顶评论**：不会漏掉 UP 主置顶

## 播报给用户的建议格式

采集完成后，简短告知：
- 视频标题 + UP
- 采到多少条评论
- Top 3 高赞评论摘要
- XLSX 文件路径 + 提醒"上传 Google Sheets 后翻译会自动执行"

用 `mcp__cowork__present_files` 把 XLSX 呈现给用户。

## 上下文

- 项目：日本 Vocaloid 音乐 IP 运营
- 核心场景：日本创作者不懂中文，需要把中国粉丝评论翻译给他们看
- 关键商业信号：评论区的"求二创授权"、"求歌词罗马音"、"求 mp3"是商业机会
