# BiliComments

> B 站视频评论收集器 · **v2 新增：AI 日语翻译（含 Vocaloid 术语表）+ 意图自动分类 + 推荐日语回复模板**。输入视频链接 → 输出增强 XLSX，日方合作曲师直接可读、可点选回复模板。

[![Python](https://img.shields.io/badge/python-3.10+-blue.svg)]()
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Claude Skill](https://img.shields.io/badge/Claude-Skill-orange.svg)]()

> 📖 **完整使用手册请看 [USAGE.md](USAGE.md)** — 包含 4 大场景化工作流（给日方发反馈报告 / 捕捉商业机会 / 短视频切片辅助 / 情感评估）、SESSDATA 详细配置、Google Sheets 完整教程、字段说明、翻译质量优化、Vocaloid 运营实战建议、FAQ、故障排查。

## 使用场景

- **日本曲师/绘师不懂中文** — 你把中文评论翻成日语给他们看，让他们知道中国粉丝在说什么
- **爆款曲评论区整理** — 抓下来做切片素材参考、评论区置顶策略
- **商业机会识别** — 「求二创授权」「求 mp3」「求歌词罗马音」类评论自动汇总
- **粉丝画像分析** — 高等级、大会员、UP 主回复自动标注

## v2 新特性 🚀

- 🇯🇵 **AI 日语翻译** — Google Translate + Vocaloid 术语表预处理（YYDS/破防/苹果p/术力口 等自动替换成日语能读懂的形式）
- 🏷️ **意图自动分类** — 强烈好评 / 情感深度 / 求二创授权 / 求资源 / 求歌词 / 负面质疑 等 12 类，每类不同颜色
- 💬 **推荐日语回复** — 每条评论 3 条日语回复模板供 UP 一键复制，负面评论额外警告"务必人工审核"
- 📊 **意图分布统计** — 视频信息页新增意图分布，一眼看到粉丝在关注什么

## 基础特性

- 🌐 **一键抓取** — 输入 bvid 或链接，输出增强 XLSX
- 🔐 **可选登录** — 匿名可用（少数热门评论），登录后可抓全量
- 🎯 **智能视图** — 3 个 Sheet：完整评论 / 视频信息 / Top 20 高赞
- 🎨 **UP 主评论高亮** — 曲师/UP 本人的评论黄色背景
- 🤖 **Claude 技能** — 附 SKILL.md，Claude 可直接调用

## 快速开始

```bash
git clone https://github.com/lepockyio-ops/bili-comments.git
cd bili-comments
pip install -r requirements.txt

# 抓热门评论（匿名即可，但只有 3-5 条）
python collect_comments.py BV1FDNJ6BE4j

# 打开 data/BV1FDNJ6BE4j.xlsx
# 上传到 Google Sheets → 翻译列自动执行
```

Windows 用户用 `run.ps1` 更方便：

```powershell
.\run.ps1 BV1FDNJ6BE4j
```

## ⚠️ 重要：关于评论数量限制

B 站 2025 后加强了反爬策略，**未登录用户只能拿到 3-5 条热门评论**。

要抓全量评论（几百到几千条），必须提供 `SESSDATA` cookie。

### 获取 SESSDATA

1. 用 Chrome/Edge 登录 https://www.bilibili.com
2. 按 `F12` 打开开发者工具
3. 切到 **Application** 标签
4. 左侧 **Cookies** → 点开 `https://www.bilibili.com`
5. 找 `SESSDATA` 那一行，复制 **Value** 列的完整字符串

### 配置 SESSDATA

**方式 A：写到 `.env` 文件（推荐）**

复制 `.env.example` 为 `.env`，填入：
```
BILI_SESSDATA=你复制的_SESSDATA_值
```

（`.env` 已在 `.gitignore` 里排除，不会误提交到 git）

**方式 B：环境变量**

Windows PowerShell:
```powershell
$env:BILI_SESSDATA = "你的_SESSDATA_值"
.\run.ps1 BV1FDNJ6BE4j
```

Linux/Mac:
```bash
export BILI_SESSDATA="你的_SESSDATA_值"
python collect_comments.py BV1FDNJ6BE4j
```

### 安全提醒

- ⚠️ SESSDATA **相当于登录凭证**，泄露 = 别人可以控制你的账号
- ⚠️ 别用主号，用一个小号专门跑爬虫
- ⚠️ 别在公共电脑设置 `.env`
- ⚠️ 别把 `.env` 提交到 git（本项目 `.gitignore` 已排除）
- ⚠️ 高频爬虫可能触发 B 站封号，建议采集间隔 ≥ 1.5 秒（本项目默认已做）

## 常用命令

```powershell
.\run.ps1 BV1FDNJ6BE4j                              # 抓热门（默认）
.\run.ps1 BV1FDNJ6BE4j --sort new                   # 按时间排序
.\run.ps1 BV1FDNJ6BE4j --sort up-only               # 只看楼主
.\run.ps1 BV1FDNJ6BE4j --max-pages 10               # 只抓前 10 页
.\run.ps1 BV1FDNJ6BE4j --include-replies            # 含所有楼中楼
.\run.ps1 BV1FDNJ6BE4j --out my.xlsx                # 自定义输出路径
.\run.ps1 BV1FDNJ6BE4j --target-lang en             # 翻成英语
.\run.ps1 https://www.bilibili.com/video/BVxxx      # 用完整链接
```

## 输出 XLSX 结构

### Sheet 1：评论

| 列 | 说明 |
|---|---|
| 楼层 | 主评论的序号 |
| 类型 | 主评论 / 楼中楼 |
| 用户名 | 评论者 B 站昵称 |
| 等级 | 用户等级 0-6 |
| 大会员 | 是 / 空 |
| **是否 UP** | 是 / 空（UP 主本人回复自动黄色高亮） |
| 勋章 | 如"黄金殿堂" |
| **评论中文（原文）** | 原始评论 |
| **翻译（ja）** | `=GOOGLETRANSLATE(...)` 公式，Google Sheets 打开自动翻译 |
| 点赞 | 该评论获赞数 |
| 回复数 | 楼中楼数量 |
| 发布时间 | CST 时区 |
| 用户签名 | 前 80 字符 |
| rpid | 评论 ID |
| 父评论 rpid | 楼中楼指向的主评论 |

### Sheet 2：视频信息

- 视频元数据（标题、UP、发布时间、时长、链接）
- 视频数据（播放/点赞/投币/收藏/评论/弹幕）
- 采集统计（时间、是否登录、条数、UP 回复数、平均等级）

### Sheet 3：Top 20 高赞

- 最热的 20 条主评论
- 含直达链接（点评论区里对应评论）
- 快速扫读用

## 打开与翻译

### 用 Google Sheets（推荐）

1. 打开 https://sheets.google.com
2. 文件 → 导入 → 上传 → 选择生成的 XLSX
3. 翻译列会自动执行 `=GOOGLETRANSLATE`，翻译结果立刻出现

### 用 Excel / WPS

翻译列会显示 `#NAME?` 错误（因为 `GOOGLETRANSLATE` 是 Google 独有函数）。

如果一定要在 Excel 里看翻译，可以：
- 手动删除翻译列的公式，改用 Excel 的 `=TRANSLATE()` 函数（Excel 365 需要 Microsoft 365 订阅）
- 或先用 Google Sheets 导入 → 复制翻译结果 → 粘贴回 Excel

## 与其他工具的联动

- **[BiliWatch](https://github.com/lepockyio-ops/biliwatch)** — 监控 UP 主投稿
- **[BiliRadar](https://github.com/lepockyio-ops/biliradar)** — 视频数据雷达
- **BiliComments（本项目）** — 评论区导出

**推荐工作流**：
1. BiliWatch 发现新曲 → 通知
2. BiliRadar 追踪数据 → 判断值不值得深耕
3. **BiliComments 导出评论 → 翻译给日方合作伙伴 → 数据支撑运营决策**

## 已知限制

| 限制 | 说明 | 缓解 |
|---|---|---|
| 匿名用户只能拿 3-5 条 | B 站反爬策略 | 配置 SESSDATA |
| 单次抓 1000+ 条可能触发风控 | IP-level 限流 | 加长采集间隔 |
| GOOGLETRANSLATE 只在 Google Sheets 生效 | Google 独有函数 | 用 Excel 时手动删公式或换函数 |
| 翻译质量取决于 Google | 二次元/日语专有名词可能翻错 | 手动修正 or 未来接 Claude API |
| 只支持视频评论 | 番剧/直播/专栏用不同接口 | v2 版本考虑加 |

## 迭代路线

- [x] v1：主评论 + 楼中楼采集 + XLSX 输出 + GOOGLETRANSLATE 公式
- [ ] v2：接 Claude API 做高质量翻译（覆盖 Google Translate 短板：二次元用语、Vocaloid 专有名词）
- [ ] v3：评论意图分类（求授权 / 求歌词 / 情感夸赞 / 负面质疑）
- [ ] v4：铁粉识别（跨曲目/跨 UP 追踪高频评论者）
- [ ] v5：弹幕采集（结合 B 站 danmaku_proto 接口）

## License

MIT © 2026 Pocky

## 免责声明

本项目使用 B 站的野生 API（非官方开放平台），仅供个人学习和轻量运营辅助。使用者应自行遵守 B 站服务条款。因滥用（如高频批量爬取、商业化数据售卖）造成的账号封禁或法律风险，与作者无关。
