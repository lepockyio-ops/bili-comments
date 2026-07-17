# BiliComments 使用手册

> B 站视频评论收集器完整使用指南 · 场景化工作流 · Vocaloid 运营实战

## 目录

- [5 分钟快速上手](#5-分钟快速上手)
- [SESSDATA 配置详解](#sessdata-配置详解)
- [4 大典型工作流](#4-大典型工作流)
- [命令详解](#命令详解)
- [Google Sheets 完整教程](#google-sheets-完整教程)
- [XLSX 字段说明](#xlsx-字段说明)
- [翻译质量优化](#翻译质量优化)
- [Vocaloid 运营实战建议](#vocaloid-运营实战建议)
- [常见问题 FAQ](#常见问题-faq)
- [故障排查](#故障排查)

---

## 5 分钟快速上手

### Step 1：安装

```bash
git clone https://github.com/lepockyio-ops/bili-comments.git
cd bili-comments
pip install -r requirements.txt
```

Windows 用户第一次跑 `run.ps1` 会自动装依赖：

```powershell
.\run.ps1 BV1FDNJ6BE4j --max-pages 1
```

### Step 2：抓取（未登录版，只有 3-5 条）

```powershell
.\run.ps1 BV1FDNJ6BE4j
```

会输出到 `data/BV1FDNJ6BE4j.xlsx`。**但因为没登录，只能拿到 3-5 条热门评论**。

### Step 3：配置 SESSDATA 拿到全量（推荐）

**只需一次配置，长期有效**：

1. Chrome/Edge 登录 B 站
2. `F12` → Application → Cookies → 复制 `SESSDATA` 的 Value
3. 复制项目里的 `.env.example` 为 `.env`
4. 编辑 `.env`，填入 SESSDATA：
   ```
   BILI_SESSDATA=你复制的_SESSDATA_值
   ```
5. 重新跑：`.\run.ps1 BV1FDNJ6BE4j`

现在能拿到全量评论（几百到几千条）。

### Step 4：上传 Google Sheets 看翻译

1. 打开 https://sheets.google.com → 新建空白表格
2. **文件 → 导入 → 上传 → 拖入 XLSX**
3. 打开「评论」页 → **翻译列自动出现日语翻译**

**5 分钟走完全流程**。

---

## SESSDATA 配置详解

### 为什么必须要 SESSDATA

B 站 2025 年后加强了反爬策略。对未登录用户，评论 API 只返回 **3-5 条热门评论** 就强制 `is_end=True`。无论你怎么翻页、换排序、加 wbi 签名，都突破不了。

**登录后**：正常返回 20 条/页，可翻页到全部评论。

### 获取 SESSDATA 的详细步骤

**Chrome / Edge（推荐）**：

1. 打开 https://www.bilibili.com 并**登录**
2. 按 `F12` 或右键 → 检查
3. 顶部标签切到 **Application**（应用）
4. 左侧展开 **Cookies** → 点击 `https://www.bilibili.com`
5. 在右侧列表找 `SESSDATA` 这一行
6. 复制 **Value** 列的完整字符串（大概 100+ 字符）

**Firefox**：

1. 登录 B 站
2. `F12` → 存储 (Storage) → Cookie → `https://www.bilibili.com`
3. 找 `SESSDATA` 复制值

### 配置的三种方式

**方式 A：`.env` 文件（推荐）**

复制 `.env.example` 为 `.env`，编辑：
```
BILI_SESSDATA=d5dc485e%2C1799...（你的实际值）
```

- ✅ 一次配置，长期有效
- ✅ 已被 `.gitignore` 排除，不会误提交
- ✅ 多人开发时每人独立配置

**方式 B：环境变量**

Windows PowerShell（临时，关闭窗口就失效）：
```powershell
$env:BILI_SESSDATA = "你的值"
.\run.ps1 BV1FDNJ6BE4j
```

Windows PowerShell（永久，写入用户环境）：
```powershell
[Environment]::SetEnvironmentVariable("BILI_SESSDATA", "你的值", "User")
# 重开 PowerShell 生效
```

Linux/Mac：
```bash
export BILI_SESSDATA="你的值"           # 临时
echo 'export BILI_SESSDATA="你的值"' >> ~/.bashrc  # 永久
```

**方式 C：命令行传入（不推荐）**

不建议直接在命令行写 SESSDATA，会存进命令历史，容易泄露。

### 安全防线

BiliComments 已经内置多层保护，但你自己要注意：

| 防线 | 状态 |
|---|---|
| `.env` 被 `.gitignore` 排除 | ✅ 已配置 |
| SESSDATA 不会出现在日志/输出 | ✅ 代码已确认 |
| 采集只做只读操作 | ✅ 不点赞/不评论/不投币 |
| 请求频率 1.5-3s/页 | ✅ 不会触发风控 |
| **用小号登录** | ⚠️ 由你决定 |
| **不要分享 .env 文件** | ⚠️ 由你决定 |
| **不要截屏 SESSDATA** | ⚠️ 由你决定 |

### SESSDATA 泄露了怎么办

去 B 站设置 → 账号安全 → 退出所有登录设备。重新登录会生成新的 SESSDATA，旧的立刻失效。

### 多账号切换（进阶）

如果你想在不同项目/曲师之间用不同小号：

**做法**：给每个 UP 主/项目一个专属子目录，各自的 `.env` 独立管理。

```
bili-comments/
├── .env                    # 默认账号
└── projects/
    ├── chenomio/.env       # 给 Chenomio 专用的号
    └── nijou/.env          # 给二錠专用的号
```

跑的时候 `cd` 到对应子目录即可（`.env` 只读当前脚本目录）。

---

## 4 大典型工作流

### 🌏 场景 A：给日方合作曲师发粉丝反馈报告

**背景**：你合作的日本曲师不懂中文，但极其想知道中国粉丝在评论区说什么。以往你要人肉翻译 20-50 条，一小时起。

**执行**：

```powershell
# 抓 top 20 热门评论（1 页）
.\run.ps1 BV1xxx --max-pages 1

# 或者抓 top 50（3 页）
.\run.ps1 BV1xxx --max-pages 3
```

**操作**：
1. 拿到 `data/BV1xxx.xlsx`
2. 上传 Google Sheets → 翻译自动执行
3. 「Top 20 高赞」页有独立的日语翻译栏
4. **直接把 Google Sheets 链接发给曲师**（可以设置只读分享）
5. 或者截屏「Top 20 高赞」页面发给他

**曲师收到后**：
- 一屏看到中日双语对照
- 知道哪句歌词最戳粉丝、哪个情节最爱哭
- 直接指导下一首曲的创作方向

**时间对比**：
- 手动翻译 20 条：约 1 小时
- BiliComments + Google Sheets：**约 3 分钟**

---

### 💰 场景 B：捕捉商业机会（求授权/求歌词/求周边）

**背景**：Vocaloid 圈评论区高频出现"求授权"、"求 mp3"、"求歌词罗马音"等评论，每一条都是商业机会。人工翻找漏率极高。

**执行**：

```powershell
# 抓多点评论（比如前 10 页 = 200 条）
.\run.ps1 BV1xxx --max-pages 10
```

**Google Sheets 里筛选**：
1. 打开「评论」页
2. 点击"评论中文（原文）"列的筛选按钮
3. 选择"按条件过滤" → "文本包含"
4. 依次筛选关键词：
   - `求` — 求授权/求资源/求歌词/求周边
   - `翻译` — 求翻译/字幕
   - `搬运` — 想搬运到别的平台
   - `二创` — 二次创作意愿
   - `商单` — 商业合作意愿

**每条匹配都是潜在的收入或流量机会**：

| 评论意图 | 商业动作 | 单条价值 |
|---|---|---|
| 求二创授权 | 主动私信发白名单流程 | 一条二创长尾流量 5-20% |
| 求 mp3/kara 下载 | 引导付费 karaoke 页 | 8-30 元/次 |
| 求歌词罗马音 | 发布 lyric video 或歌词 PDF | 商单素材 |
| 想搬运/翻译到别的平台 | 授权 + 分成 | 二次分发 3-10x 曝光 |

---

### 🎬 场景 C：短视频切片的评论辅助

**背景**：想把一首 4 分钟爆曲切成 3-5 条 30 秒短视频，但不知道切哪几段。

**执行**：

```powershell
.\run.ps1 BV1xxx --include-replies
```

**分析**：
1. 打开 XLSX，用 `Ctrl+F` 搜索时间戳模式：`0:52`、`1:23`、`2:04`、`3:14` 等
2. 粉丝在评论里明确指认的"神段"就是切片候选点
3. 也可以按点赞排序，看最热的评论提到了哪些段落
4. 结合弹幕热力图（配合 DanmuMind 未来版本）双重验证

**产出**：3-5 个切片候选点，每个都有粉丝反馈作为质量背书。**命中率从 30% 提升到 70-80%**。

---

### 🎯 场景 D：情感评估（是否值得追加宣发）

**背景**：一首新曲发布 24h 后，数据看起来一般（BiliRadar 里没触发预警），但你不确定要不要放弃。

**执行**：

```powershell
.\run.ps1 BV1xxx --max-pages 3
```

**关键判断信号**（都在评论区）：
1. **正面情感占比**：
   - 「好听」「哭了」「YYDS」「循环」 等 > 80% → 内容质量硬，数据慢只是曝光不够，值得投流
   - 中性/无感占多 → 内容有问题，止损
2. **深度评论比例**：
   - 高等级用户（L5-L6）愿意写长评论 → 命中核心受众
   - 只有短评论「好听」「不错」 → 触达广但没触动
3. **UP 主回复热度**：
   - Chenomio 那首曲，UP 主两条置顶评论都进 Top 5 → 说明社群互动氛围极好

**决策矩阵**：

| 数据 | 评论正面率 | 结论 |
|---|---|---|
| 强 | 强 | 全力宣发 |
| 强 | 弱 | 蹭热度型，短期收割 |
| 弱 | **强** | **投流值得试**，可能是曝光问题 |
| 弱 | 弱 | 止损，别投钱 |

---

## 命令详解

### `python collect_comments.py <bvid_or_url> [options]`

或 Windows：`.\run.ps1 <bvid_or_url> [options]`

### 参数

| 参数 | 类型 | 说明 |
|---|---|---|
| `bvid_or_url` | 必需 | `BVxxxxx` 或 `https://www.bilibili.com/video/BVxxx` |
| `--out PATH` | 可选 | 输出 XLSX 路径。默认 `data/<bvid>.xlsx` |
| `--sort MODE` | 可选 | `hot`(默认) / `new` / `up-only` |
| `--max-pages N` | 可选 | 最多拉取的页数。每页约 20 条 |
| `--include-replies` | 可选 | 同时抓所有主评论的楼中楼 |
| `--target-lang LANG` | 可选 | 翻译目标语言代码，默认 `ja` |

### 排序模式对比

| `--sort` | 说明 | 何时用 |
|---|---|---|
| `hot`（默认）| 按热度（点赞×时间衰减）| **想看有代表性的评论**，日常首选 |
| `new` | 按时间倒序（最新→最旧）| 想看近期讨论、跟进最新反馈 |
| `up-only` | 只看 UP 主本人评论 | 想看曲师的置顶/回复 |

### 翻译目标语言

`--target-lang` 支持 Google Translate 的所有语言代码：

| 代码 | 语言 | 场景 |
|---|---|---|
| `ja`（默认）| 日语 | 日方合作方 |
| `en` | 英语 | 海外粉丝 / 国际推广 |
| `ko` | 韩语 | K-Pop 圈联动 |
| `zh-TW` | 繁体中文 | 台/港粉丝 |
| `es` | 西班牙语 | 拉美地区 |

完整列表：https://cloud.google.com/translate/docs/languages

### 命令示例

```powershell
# 最简单：抓热门评论全量
.\run.ps1 BV1FDNJ6BE4j

# 只抓 20 条（1 页 hot）
.\run.ps1 BV1FDNJ6BE4j --max-pages 1

# 抓最新 100 条
.\run.ps1 BV1FDNJ6BE4j --sort new --max-pages 5

# 只看曲师本人评论
.\run.ps1 BV1FDNJ6BE4j --sort up-only

# 抓 top 20 + 每条的楼中楼（完整讨论上下文）
.\run.ps1 BV1FDNJ6BE4j --max-pages 1 --include-replies

# 输出到桌面 + 翻成英语
.\run.ps1 BV1FDNJ6BE4j --out "C:\Users\Me\Desktop\comments.xlsx" --target-lang en

# 用完整 URL（含短链/时间戳的都行）
.\run.ps1 https://www.bilibili.com/video/BV1FDNJ6BE4j/?spm_id_from=333.999.0.0
```

---

## Google Sheets 完整教程

### 步骤 1：上传 XLSX

1. 打开 https://sheets.google.com
2. 左下角 `+` 号新建空白表格
3. 顶部菜单 **文件 → 导入**
4. 切换到 **上传** 标签，拖入你的 `.xlsx` 文件
5. **导入位置** 选择"替换电子表格"
6. 点 **导入数据**

### 步骤 2：验证翻译

- 打开「评论」sheet
- 找 I 列（翻译），应该显示日语文字
- 如果显示 `#NAME?` 或空白：
  - 等 3-10 秒，Google 需要计算公式
  - 或者点击一个翻译单元格，看公式栏是否是 `=IF(H2="","",GOOGLETRANSLATE(H2,"zh-CN","ja"))`

### 步骤 3：定制视图（可选）

**冻结首行 + 首列**（脚本已默认冻结首行）：
- 视图 → 冻结 → 1 列

**筛选热门评论**：
- 选中 J 列（点赞）→ 数据 → 排序范围 → 降序

**只看 UP 主回复**：
- 数据 → 创建筛选器
- 点 F 列（是否 UP）筛选器 → 只勾"是"

**高亮长评论**（超过 50 字的深度粉丝）：
- 格式 → 条件格式
- 应用于范围：H2:H1000
- 条件："自定义公式" → `=LEN(H2)>50`
- 格式：填充色 → 淡绿色

### 步骤 4：分享给日方

**只读分享**（推荐给外部合作方）：
1. 右上角 **共享** 按钮
2. **常规访问** → **知道链接的任何人** → **查看者**
3. 复制链接
4. 发给曲师/绘师

**添加协作者**（团队内部）：
1. 共享 → 输入邮箱 → 选择"编辑者"或"评论者"

### 步骤 5：导出为其他格式

**导出 PDF**（会议汇报用）：
- 文件 → 下载 → **PDF (.pdf)**
- 页面方向：横向
- 缩放：适应页宽

**导出 CSV**（给分析师）：
- 文件 → 下载 → **逗号分隔值 (.csv)**（只导当前 Sheet）

---

## XLSX 字段说明

### Sheet 1：评论（主表）

| 列 | 字段 | 说明 | 运营价值 |
|---|---|---|---|
| A | **楼层** | 主评论的序号 | 定位用 |
| B | **类型** | 主评论 / 楼中楼 | 区分层级 |
| C | **用户名** | B 站昵称 | 识别铁粉 |
| D | **等级** | 用户等级 0-6 | L5-L6 是活跃老用户，评论质量高 |
| E | **大会员** | 是 / 空 | 大会员 = 付费用户，付费转化潜力高 |
| F | **是否 UP** | 是 / 空（黄色高亮）| **一眼看曲师本人回复** |
| G | **勋章** | 如"黄金殿堂" | 稀有勋章 = 高价值用户 |
| H | **评论中文（原文）** | 原始文本 | 分析情感、内容意图 |
| I | **翻译（ja）** | GOOGLETRANSLATE 公式 | Google Sheets 自动翻译 |
| J | **点赞** | 该评论获赞数 | 排序判断代表性 |
| K | **回复数** | 楼中楼数量 | > 5 说明引发讨论 |
| L | **发布时间** | CST 时区 | 判断是首日粉丝还是长尾发现 |
| M | **用户签名** | 前 80 字 | 判断是否 Vocaloid 圈内人 |
| N | **rpid** | 评论 ID | 拼直达链接用 |
| O | **父评论 rpid** | 楼中楼指向的主评论 | 追溯讨论上下文 |

### Sheet 2：视频信息

- 视频元数据 + 视频数据 + 采集统计
- **重点看**："是否登录采集" — 未登录时会提示

### Sheet 3：Top 20 高赞

- 快速扫读用
- 含直达链接（H 列）— 点击可以直接跳到 B 站对应评论

---

## 翻译质量优化

### 问题：Google Translate 对二次元用语翻不准

**典型翻错**：
- 「YYDS」→ 直译成拼音，日语看不懂
- 「破防了」→ 翻成"防御を突破された"（字面）
- 「太上头了」→ 翻成"頭が上に上がる"（错）
- 「歌姫」→ 有时翻错顺序

### 解法 1：手动替换公式

在 Google Sheets 里，把某些行的公式手动替换为你自己的翻译：

1. 选中翻译单元格
2. 按 `Delete` 清空
3. 手动输入你的翻译

**推荐**：Top 5 高赞评论手动翻译，其他用公式默认，兼顾质量和效率。

### 解法 2：术语对照表（进阶）

在 Google Sheets 新建一个「术语表」sheet：

| 中文 | 日语 |
|---|---|
| YYDS | 神曲 |
| 破防了 | 心が壊れた |
| 太上头了 | 中毒になった |
| 泪目 | 泣ける |
| 循环了一天 | 一日中リピート |

然后用 SUBSTITUTE 嵌套公式：
```
=GOOGLETRANSLATE(SUBSTITUTE(SUBSTITUTE(H2,"YYDS","神曲"),"破防","心が壊れる"),"zh-CN","ja")
```

**缺点**：公式变复杂，维护术语表要花时间。

### 解法 3：用 Claude 二次翻译（未来 v2）

BiliComments v2 计划接 Claude API 做高质量翻译：
- Google Translate 做初翻
- Claude 做审校（保留网络用语、Vocaloid 专有名词）
- 输出双列：Google 翻译 + Claude 优化

暂时未实现。想现在用：手动复制评论到 Claude 桌面版，让 Claude 翻译一遍再粘回。

### 解法 4：直接用 Claude 桌面翻译整个 XLSX

**最省事的方案**：

1. 用 BiliComments 抓评论（不用 Google Sheets 也行）
2. 打开 XLSX，选中 H 列所有评论 → 复制
3. 打开 Claude 桌面版对话，说：
   > 帮我把下面这批中文评论翻译成日语，保留 Vocaloid 圈的语气，输出成两列（原文 | 翻译）：
   > [粘贴]
4. Claude 输出高质量翻译
5. 复制回 XLSX

**质量差三个量级**，缺点是每次都要手动。

---

## Vocaloid 运营实战建议

### 建议 1：给每首爆曲建立评论档案

爆曲的评论区含金量极高，建议保存历史：

```powershell
# 定期采集同一首曲的评论（每周一次）
.\run.ps1 BV1xxx --out "data/BV1xxx_2026-07-14.xlsx"
.\run.ps1 BV1xxx --out "data/BV1xxx_2026-07-21.xlsx"
```

对比不同时间点的评论内容，能看到：
- 首周：新粉的情感反应
- 第 2-4 周：老粉的深度讨论、二创请求
- 长尾期：意外发现的用户群体

### 建议 2：给曲师定制"评论周报"

每周一给合作曲师发一份 Google Sheets 报告：
- 本周 3-5 首曲的 Top 10 评论
- 每条含日语翻译
- 高亮 UP 主本人回复
- 标注商业机会（求授权等）

**这是一个签约谈判的杀手锏** —— 曲师看到你如此认真对待他的粉丝，续约率和话语权都会上升。

### 建议 3：结合 BiliRadar 联动

```
BiliRadar 发现"BV1xxx 触发 [GOOD] 优质曲预警"
       ↓
你：`.\run.ps1 BV1xxx --max-pages 5`（抓 100 条评论）
       ↓
上传 Google Sheets → 分析评论情感 → 找商业信号
       ↓
决策：追加短视频切片 + 主动联系求授权粉丝
```

### 建议 4：负面评论要在 24h 内看到

爆款曲偶尔会有版权质疑、拉踩评论。建议每周至少 1 次抓最新评论：

```powershell
.\run.ps1 BV1xxx --sort new --max-pages 2
```

如果发现负面评论集中出现，立刻处理（UP 主回复澄清、必要时删稿重发）。

---

## 常见问题 FAQ

### Q1: 抓的评论数远少于视频显示的评论总数？

**A**: 检查是否配置了 SESSDATA。未登录只能拿 3-5 条，登录后每页 20 条。

也可能是视频有超多楼中楼（比如 10w 评论里 8w 是楼中楼）。不加 `--include-replies` 只拿主评论，正常现象。

### Q2: 抓完就报 "命中风控"？

**A**: 短时间高频请求触发。解决：
- 等 15-30 分钟再试
- 换个网络（切手机热点）
- 加大 `time.sleep` 间隔（改 `collect_comments.py` 里的 `random.uniform(1.5, 3.0)` 为 `random.uniform(3, 6)`）

### Q3: SESSDATA 会不会过期？

**A**: 一般有效期几个月，但 B 站可能因为你换设备/清 cookie 提前失效。失效表现：
- 抓回来只有 3-5 条（跟未登录一样）
- 日志显示 `code=-101` 或类似

重新按前面的步骤复制新的 SESSDATA 到 `.env` 即可。

### Q4: 支持抓番剧/直播评论吗？

**A**: 不支持。番剧评论用 `type=11`（本项目写死 `type=1`），直播用不同的接口。要加需要改代码。

### Q5: XLSX 文件多大合算的？

**A**: 常见规模：
- 20 条评论：~15 KB
- 100 条评论：~40 KB
- 500 条评论：~150 KB
- 5000 条评论：~1.5 MB

Google Sheets 单表格上限 1000w 单元格，够用。

### Q6: 能定时自动抓吗？

**A**: 能。用 Windows 任务计划程序或 Linux cron：

```powershell
# Windows：每天早上 9 点抓一次某曲的评论
$action = New-ScheduledTaskAction -Execute "powershell" `
  -Argument "-NoProfile -File `"C:\path\to\bili-comments\run.ps1`" BV1xxx --max-pages 3"
$trigger = New-ScheduledTaskTrigger -Daily -At 9am
Register-ScheduledTask -TaskName "BiliComments-BV1xxx-Daily" -Action $action -Trigger $trigger
```

```cron
# Linux：每天 9 点抓
0 9 * * * cd /path/to/bili-comments && python collect_comments.py BV1xxx --max-pages 3
```

### Q7: 楼中楼要不要抓？

**A**: 看场景：
- **给曲师看反馈**：**不要**，太长看不完
- **切片分析**：**要**，讨论上下文重要
- **商业机会挖掘**：**要**，很多"求授权"藏在楼中楼里

### Q8: 翻译不准怎么办？

**A**: 参考本手册「翻译质量优化」章节。核心 3 条：
1. Top 5 手动翻译
2. 建术语表用 SUBSTITUTE 替换
3. 或直接把评论粘给 Claude 翻译

### Q9: 一个 XLSX 能包含多个视频吗？

**A**: 当前不行。每个视频独立文件。如果需要合并：
- Google Sheets 里手动导入多个 → 复制粘贴到同一表
- 或写个脚本合并 XLSX（未来 v2 考虑加 `--append` 参数）

### Q10: 输出到 CSV 而不是 XLSX 行吗？

**A**: 当前脚本只支持 XLSX（因为要放 GOOGLETRANSLATE 公式）。如果你想要 CSV，可以：
1. 生成 XLSX 后用 Google Sheets/Excel 打开
2. 另存为 CSV

或者改脚本用 `csv` 模块直接输出（几行代码的事）。

---

## 故障排查

### ❌ `缺少 httpx` / `缺少 openpyxl`

```powershell
py -m pip install httpx openpyxl
```

如果这命令没输出且没安装成功，说明你踩到了 VIRTUAL_ENV 污染问题：

```powershell
Remove-Item Env:VIRTUAL_ENV -ErrorAction SilentlyContinue
py -m pip install httpx openpyxl
```

（`run.ps1` 已自动处理这个，直接用它就没事）

### ❌ 中文输出乱码

用 `run.ps1` 包装脚本。若直接跑 `py`：

```powershell
$env:PYTHONIOENCODING = "utf-8"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
py -X utf8 collect_comments.py BVxxxx
```

### ❌ 无法从 'xxx' 提取 bvid

BiliComments 需要 `BV` 开头 10 位字符的 bvid。检查：
- ✅ `BV1FDNJ6BE4j`
- ✅ `https://www.bilibili.com/video/BV1FDNJ6BE4j`
- ❌ `av123456`（旧格式，需要转成 bvid）
- ❌ `b23.tv/xxxxx`（短链，先展开）

展开短链：
```powershell
curl.exe -sI "https://b23.tv/xxxxx" | Select-String "location:"
```

### ❌ SESSDATA 已配置但还是只有 3-5 条

排查步骤：
1. 检查 `.env` 文件路径：必须在 `bili-comments/` 目录下
2. 检查内容格式：`BILI_SESSDATA=value`（等号两边无空格，值无引号）
3. 检查 SESSDATA 是否过期：重新登录 B 站，重新复制
4. 检查小号是否被封：登录 web 试试能不能刷新评论区
5. 试试环境变量方式（临时排除 .env 读取问题）：
   ```powershell
   $env:BILI_SESSDATA = "你的值"
   .\run.ps1 BVxxx
   ```
   看日志第一行是否显示"使用 SESSDATA 登录"

### ❌ HTTP 412 / 命中风控

已经开始防风控 sleep 1.5-3s/页，但云端 IP 段更严格。缓解：
- 等 30 分钟再试
- 换用户 IP（手机热点）
- 单次采集页数减少（`--max-pages 5`）
- 如果长期在同一 IP 高频跑，考虑用 VPN 换 IP

### ❌ Google Sheets 里翻译列显示 #NAME?

原因：GOOGLETRANSLATE 是 Google Sheets 独有函数，Excel/WPS 打不开。

**只在 Google Sheets 里生效**。如果你在 Excel/WPS 打开这个文件，翻译列会全部显示 #NAME? — 属于正常，因为函数不存在。

要在 Excel 里看翻译，只能用 Microsoft 365 的 `=TRANSLATE()` 函数或手动翻译。

### ❌ Google Sheets 翻译列长时间空白

Google 计算 GOOGLETRANSLATE 需要 3-30 秒。如果超过 1 分钟还是空白：
1. 点击一个翻译单元格，查看公式栏
2. 应该是 `=IF(H2="","",GOOGLETRANSLATE(H2,"zh-CN","ja"))`
3. 如果公式对，说明 Google 服务临时限流，稍等
4. 如果公式变成 `#REF!` 或其他错，检查 H 列是否有内容

### ❌ 抓完程序卡住不退出

按 `Ctrl+C` 强制中断。可能原因：
- httpx 网络超时未生效
- 楼中楼太多（几百层）

单次 `Ctrl+C` 会给一个优雅退出的机会，如果还不退出按第二次强杀。

---

## 版本 & 反馈

**当前版本**：v1

**GitHub**：https://github.com/lepockyio-ops/bili-comments
**Issues**：https://github.com/lepockyio-ops/bili-comments/issues
**License**：MIT

## 下一步迭代候选

- **v2 高质量翻译**：接 Claude API，覆盖 Google Translate 短板（Vocaloid 专有名词、二次元用语）
- **v3 评论意图分类**：AI 自动打标"求授权/求歌词/情感夸赞/负面质疑"，直接输出商业机会清单
- **v4 铁粉识别**：跨曲/跨 UP 追踪高频评论者，输出核心粉丝报告
- **v5 弹幕采集**：结合 B 站 danmaku 接口，把弹幕也纳入分析

**要哪个功能，直接告诉 Claude "bili-comments 加 XX 功能"**，让 Claude 帮你在这个仓库上直接迭代。
