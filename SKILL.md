---
slug: wechat-oa
displayName: 微信公众号草稿箱管理工具
summary: 微信公众号草稿箱管理工具集，支持创建/更新/删除草稿、上传素材、生成封面图，基于官方API，无需第三方依赖。
license: MIT
name: wechat-oa
description: WeChat Official Account draft management toolkit. Trigger words: 看看草稿箱/查看草稿/草稿列表/公众号草稿/搜草稿/搜索草稿/创建草稿/新建草稿/发文章到公众号/推送文章/更新草稿/删除草稿/生成封面图/上传图片/生成配图. Official API, no third-party dependencies.
description_zh: 微信公众号草稿箱管理工具集。触发词（满足任一即触发）：看看草稿箱/查看草稿/草稿列表/公众号草稿/搜草稿/搜索草稿/按关键词找草稿/按标题搜/创建草稿/新建草稿/发文章到公众号/推送文章/更新草稿/删除草稿/批量删除草稿/生成封面图/上传图片到公众号/上传图片到素材库/已发布文章列表/公众号素材列表/素材管理/删除素材/交互式删除/批量删除素材/关键词过滤素材/生成配图/生成信息图/去AI味/去Al味/文字改写/quaiwei。官方API，无需第三方依赖。
version: "2.0.5"
author: Woody
email: andy8663@163.com
wechat_mp: 用技术定义未来
homepage: https://github.com/andy8663/wechat-oa
metadata:
  openclaw:
    emoji: "📝"
    category: "publishing"
    requires:
      bins: ["python3"]
    voice_commands:
      - "查看草稿箱"
      - "看看公众号草稿"
      - "创建一篇公众号文章"
      - "推送文章到公众号"
      - "生成封面图"
      - "上传图片到素材库"
      - "生成正文配图"
      - "去AI味"
      - "文字改写"
      - "去除AI痕迹"
---

# wechat-oa

微信公众号草稿箱管理工具集。基于官方微信 API，无需第三方依赖。

WeChat Official Account draft management toolkit. Built on official WeChat APIs, no third-party dependencies required.

---

## ⚠️ 使用前必读 / MUST READ BEFORE USE

**创建或更新公众号文章前，AI 必须先阅读以下两个文档：**

Before creating or updating any WeChat article, AI MUST read these two documents first:

### 1. `WECHAT_STYLING.md`（微信渲染兼容规范）

> **为什么需要这个文档？** wechat-oa 的推送管道会删除所有块级元素的 `margin` 属性（`_clean_wechat_margins()`），导致段落间距消失。本文档提供解决方案。

Why this document? wechat-oa's push pipeline removes all `margin` attributes from block elements, causing paragraph spacing to disappear. This document provides solutions.

**核心规则 / Key rules:**
- 段落间距用 `<br>` 实体换行，不用 CSS `margin`/`padding`
- h2：17px、居中、`#2563eb`、前面加 `<br>`
- h3：15px、`#b91c1c`（与蓝色互补）
- `<br>` 只插在顶级块之间，不插在列表/blockquote/提示框内
- CSS 禁止重复属性（后面的覆盖前面的）
- footer：`#999` 灰色、居中、前面加 `<br>`

📖 **完整规范见 `WECHAT_STYLING.md`**

### 2. `design.md`（整体排版风格）

> 定义公众号文章的整体视觉风格。

Defines the overall visual style of WeChat articles.

**涵盖内容 / Covers:**
- 容器宽度（文章 677px / 图文卡片 375px）
- 字体规范（clamp() 响应式字号）
- 配色规范（≤5 主体色，对比度 ≥4.5:1）
- 布局规范（Flex/Grid only，禁止 absolute/float）
- 标题规范（禁止重复主标题、header 标签、左边框装饰）
- 内容结构（扁平结构，有/无背景色的 padding 规则不同）
- CSS/HTML/JS 限制（公众号渲染器兼容性白名单）

📖 **完整规范见 `design.md`**

---

## 🚀 快速开始 / Quick Start

### 1. 安装依赖 / Install Dependencies

```bash
pip install requests Pillow
```

### 2. 配置凭证 / Configure Credentials

```bash
cp config.example.json config.json
# 编辑 config.json，填入 APP_ID 和 APP_SECRET
# Edit config.json, fill in APP_ID and APP_SECRET
```

### 3. 试试看 / Try It Out

```bash
# 查看草稿列表
python wechat_push.py list

# 创建新草稿
python wechat_push.py create article.html
```

---

## 📦 安装与配置 / Installation & Configuration

### 获取 AppID 和 AppSecret / Get AppID & AppSecret

1. 登录 [微信公众平台](https://mp.weixin.qq.com)
2. 进入 **设置与开发 → 基本设置**
3. 复制 **公众号 AppID** 和 **公众号 AppSecret**（如未设置需先启用）

### IP 白名单配置 / IP Whitelist Configuration

调用微信 API 前，必须将服务器 IP 加入白名单：

| 推送模式 | 需要加入白名单的 IP | 说明 |
|---------|------------------|------|
| `direct` | 本机出口 IP | 直连微信 API，配本机 IP |
| `hybrid` | 本机出口 IP | 优先直连，失败自动切中转 |
| `relay` | **服务器 IP** `120.79.2.44` | 通过中转服务器调用微信 API |

**配置步骤：**
1. 登录微信公众平台
2. 进入 **设置与开发 → 安全中心 → IP 白名单 → 点击「配置」**
3. 填入对应模式的 IP（多个 IP 用回车分隔）
4. 保存

> ⚠️ **不添加 IP 白名单会导致 API 调用失败！**

💡 **不方便配本机 IP 白名单？** 使用 `hybrid` 或 `relay` 模式，只需将服务器 IP `120.79.2.44` 加入白名单即可。

### 配置文件 / Config File

将 `config.example.json` 复制为 `config.json`，填入你的凭证：

```json
{
  "APP_ID": "wx0000000000000000",
  "APP_SECRET": "00000000000000000000000000000000",
  "author": "龙虾",
  "PUSH_MODE": "hybrid",
  "WECHAT_OA_SERVER": "http://120.79.2.44",
  "WECHAT_OA_SERVER_KEY": ""
}
```

`PUSH_MODE` 说明：

| 模式 | 说明 | 适用场景 |
|------|------|----------|
| `direct` | 直连微信官方 API | IP 固定且可配白名单 |
| `relay` | 通过中转服务器推送 | IP 不固定，无法配白名单 |
| `hybrid` | 优先直连，失败自动切换中转 | **推荐** — 兼顾速度与稳定性 |

> 🔑 **如何获取 WECHAT_OA_SERVER_KEY？**  
> 发送邮件到 `andy8663@126.com` 申请中转服务器 API Key。

> ⚠️ `config.json` 包含凭证，**不要提交到 GitHub**！已在 `.gitignore` 中忽略。

### 环境切换 / Environment Switching

客户端支持 prod（生产）和 dev（开发）两个服务端环境：

| 方式 | 说明 |
|------|------|
| `config.json` 的 `ENV` 字段 | 不写或 `"ENV": "prod"` → `/api/`（生产，默认）<br>`"ENV": "dev"` → `/testapi/`（开发） |
| `--env dev\|prod` 参数 | 临时覆盖，不修改配置文件 |

```bash
# 默认连生产（prod）
python wechat_push.py create article.html

# 临时切换到开发环境
python wechat_push.py create article.html --env dev
```

**开发流程：**
1. 改代码 → 推 `dev` 分支 → 用 `--env dev` 测试
2. 测试通过 → 合入 `main` 分支
3. 部署生产 → 用 `--env prod` 或不传（默认）

---

## 📝 核心功能 / Core Features

### 草稿管理 / Draft Management

| 命令 | 说明 | 示例 |
|------|------|------|
| `list` | 查看草稿列表 | `python wechat_push.py list` |
| `find <关键词>` | 按标题搜索草稿 | `python wechat_push.py find "AI"` |
| `create <文件>` | 创建新草稿（支持 .html 和 .md） | `python wechat_push.py create article.html` |
| `update <media_id> <文件>` | 更新已有草稿 | `python wechat_push.py update <media_id> article.html` |
| `delete <media_id>` | 删除草稿 | `python wechat_push.py delete <media_id>` |
| `batch-del <id1> [id2] ...` | 批量删除草稿 | `python wechat_push.py batch-del <id1> <id2>` |

### 素材管理 / Material Management

| 命令 | 说明 | 示例 |
|------|------|------|
| `upload <图片>` | 上传图片到永久素材库 | `python wechat_push.py upload cover.png` |
| `materialcount` | 获取素材总数 | `python wechat_push.py materialcount` |
| `materials [type] [count]` | 批量获取素材列表 | `python wechat_push.py materials image 10` |
| `materialdel [media_id...]` | 批量删除素材 | `python wechat_push.py materialdel <media_id>` |
| `published` | 获取已发布文章列表 | `python wechat_push.py published` |

### 自动功能 / Auto Features

- **自动封面生成**：根据文章标题 AI 生成科技风封面图（2.35:1 比例）
- **正文图片自动上传**：自动提取 HTML/MD 中的本地图片，上传到微信素材库并替换 URL
- **智能摘要**：AI 推送文章时生成 1-2 句精准摘要传入 `digest` 参数

---

## 🎨 高级功能 / Advanced Features

### 中继模式 (Relay Mode) / AI 收支付

当 `config.json` 中 `PUSH_MODE` 设为 `relay` 时，文章通过公网服务器（wechat-oa-server）中转推送到微信公众号。中继模式支持 **支付宝 AI 收** 标准协议（HTTP 402 + Payment-Needed）。

**推送流程（免费模式）：**
```bash
python wechat_push.py create article.html
```

**推送流程（收费模式）：**
1. 调用 push_article → 服务端返回 HTTP 402 + Payment-Needed
2. 客户端保存 Payment-Needed → 调用 `alipay-bot` 发起支付
3. 用户扫码完成支付 → 告诉 Agent "已支付"
4. 调用 `finish_push(trade_no, payload)` → 服务端验证 → 执行推送

### 摘要（digest）生成规范 / Digest Generation

**推送或更新文章时，AI 必须生成摘要并传入 `digest` 参数，不要留空让服务端自动提取。**

**摘要要求：**

| 维度 | 规范 |
|------|------|
| 长度 | 1-2 句话，不超过 120 字（微信限制 128 字，留余量） |
| 内容 | 概括文章核心观点或亮点，不是机械截取正文前几句 |
| 风格 | 简洁有吸引力，让读者在公众号消息列表中有点击欲望 |
| 语言 | 与文章正文语言一致 |

### 正文配图自动生成 / Auto Infographic Generation

`generate_infographic.py` 可以根据章节内容自动生成配图，无需 AI API，完全本地 PIL 生成。

**支持的配图类型：**

| 类型 | 用途 | 示例 |
|------|------|------|
| `steps` | 流程图 | 汇款步骤：注册→填表→汇款→完成 |
| `comparison` | 对比图 | 传统汇款 vs 西联汇款 |
| `timeline` | 时间线 | 2020→2022→2024 发展历程 |
| `textcard` | 文字卡片 | 金句、要点提炼 |
| `stats` | 数据统计图 | 各渠道手续费对比 |

**使用方式：**
```bash
python generate_infographic.py steps output/step.png "步骤1" "步骤2" "步骤3"
python generate_infographic.py comparison output/compare.png "优点:很好用" "缺点:有点贵"
```

### 去AI味 / Quaiwei

去除文字中的AI生成痕迹（按次收费1元）：

```bash
python wechat_push.py quaiwei "这款产品值得注意的是，综上所述，此外还有很好的用户体验..."
```

---

## 📚 命令参考 / Command Reference

| 命令 Command | 说明 Description | 底层API Underlying API |
|------|------|---------|
| `list` | 查看草稿列表（含标题+时间） | `draft/batchget` |
| `find <关键词>` | 按标题关键词搜索草稿 | `draft/batchget` |
| `get <media_id> [--save]` | 获取单篇草稿详情（--save 保存 HTML） | `draft/get` |
| `create <文件> [--env dev\|prod]` | 创建新草稿（支持 .html 和 .md） | `draft/add` |
| `update <media_id> <文件> [--env dev\|prod]` | 更新已有草稿 | `draft/update` |
| `update <media_id> <文件> --force-cover` | 更新草稿并强制重新生成封面 | `draft/update` |
| `delete <media_id>` | 删除草稿 | `draft/delete` |
| `batch-del <id1> [id2] ...` | 批量删除草稿 | `draft/delete` |
| `upload <图片文件>` | 上传图片到永久素材库 | `material/add_material` |
| `materialcount` | 获取各类永久素材总数 | `material/get_materialcount` |
| `materials [type] [count] [offset] [keyword]` | 批量获取素材列表 | `material/batchget_material` |
| `materialdel [media_id...]` | 批量删除素材 | `material/del_material` |
| `published` | 获取已发布文章列表 | `material/batchget_material` |
| `cover <标题>` | 生成封面图预览（不推送） | PIL local generation |
| `infographic <类型> <输出路径> [参数]` | 生成正文配图 | PIL local generation |
| `quaiwei <文字内容>` | 去AI味 | Claude API + 支付宝AI收 |

---

## 🖼️ 正文图片处理 / Inline Image Processing

### 自动上传流程 / Auto-upload Flow

创建或更新草稿时，系统会自动处理正文中的图片：

```
HTML/MD 文件
    ↓
提取 <img src="..."> 或 ![alt](path)
    ↓
本地图片？ ──是──→ 上传到微信素材库 ──→ 获取微信 URL
    ↓ 否
网络图片？ ──是──→ 保留原 URL（可选下载上传）
    ↓ 否
跳过（已是微信素材库图片）
    ↓
替换 HTML 中的 src 为微信 URL
    ↓
推送草稿
```

### 支持的图片格式 / Supported Formats

- **HTML**: `<img src="local/image.png">` 或 `<img src="http://example.com/img.png">`
- **Markdown**: `![描述](./images/photo.jpg)`
- **路径类型**: 相对路径、绝对路径、`file:///` 协议

---

## ❓ 常见问题 / FAQ

### Q1：为什么我的文章段落之间没间距？

A：`_clean_wechat_margins()` 删除了所有 `margin`。解决方法：用 `<br>` 代替，详见 `WECHAT_STYLING.md`。

### Q2：为什么标题颜色没生效？

A：可能 CSS 里 `color` 属性出现了两次，后面的覆盖前面的。检查 `WECHAT_STYLING.md` 的"常见问题"章节。

### Q3：推送时提示 IP 不在白名单？

A：按照"安装与配置 → IP 白名单配置"章节，将对应 IP 加入白名单。

### Q4：如何切换到开发环境测试？

A：使用 `--env dev` 参数，或修改 `config.json` 添加 `"ENV": "dev"`。

---

## 💡 特性 / Features

- **公众号排版规范** / WeChat MP layout specification：内置 `design.md` 和 `WECHAT_STYLING.md` 排版规范，AI 生成 HTML 时必须遵循
- **行内样式转换** / Inline-style conversion：自动将 HTML 中的 `<style>` 标签转换为行内 `style=""` 属性，兼容微信文章渲染
- **自动封面生成** / Auto cover generation：根据文章标题 AI 生成科技风封面图（2.35:1 比例）
- **正文图片自动上传** / Auto inline image upload：自动提取 HTML/MD 中的本地图片，上传到微信素材库并替换 URL
- **智能摘要** / Smart digest：AI 推送文章时生成 1-2 句精准摘要

---

## 📝 输出文件 / Output Files

- `draft_ids.txt` - 草稿记录（创建时间、标题、media_id）
- 封面图默认保存在 HTML 文件同目录

---

## 🐛 问题反馈 / Feedback

- 🌐 **GitHub Issues**：[https://github.com/andy8663/wechat-oa](https://github.com/andy8663/wechat-oa)
- 📧 **邮箱**：`andy8663@126.com`
- 🔔 **微信公众号**：技术定义未来（ID: `gh_b906288c4c2f`）

> 💡 提交 Issue 前建议先搜索是否已有类似问题

---

## 📖 相关文档 / Related Documents

- **`WECHAT_STYLING.md`** - 微信渲染兼容规范（必读）
- **`design.md`** - 整体排版风格规范（必读）
- **`README.md`** - 项目介绍与安装指南
