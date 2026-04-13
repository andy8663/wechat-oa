# wechat-oa

**微信公众号「技术定义未来」官方工具** · 基于官方微信 API，无需第三方依赖

English: A WeChat Official Account draft management toolkit based on the official WeChat API. No third-party dependencies required.

---

## 功能 Features

- 📋 查看草稿列表 / List drafts
- ✏️ 创建新草稿（自动生成封面图） / Create new drafts (auto-generates cover)
- 🔄 更新已有草稿 / Update existing drafts
- 🗑️ 删除草稿 / Delete drafts
- 🖼️ 独立生成封面图预览 / Generate cover image preview independently

---

## 配套公众号 Companion Account

本工具由公众号 **「技术定义未来」** 配套开发，用于高效管理公众号草稿箱。

- 📢 公众号：技术定义未来
- 🔍 定位：AI工具效率提升 + AI商业化场景探索
- 🔗 关注方式：微信搜索「技术定义未来」或扫描下方二维码

---

## 快速开始 Getting Started

### 1. 安装依赖 Install dependencies

```bash
pip install requests Pillow
```

### 2. 配置凭证 Configure credentials

```bash
cp config.example.json config.json
# 编辑 config.json，填入你的 AppID 和 AppSecret
# Edit config.json, fill in your AppID and AppSecret
```

config.json 模板：

```json
{
  "APP_ID": "",
  "APP_SECRET": "",
  "author": ""
}
```

> **注意 / Note**: 请前往 [微信公众平台](https://mp.weixin.qq.com) → 设置与开发 → 基本配置，获取 AppID 和 AppSecret。
> Get your AppID and AppSecret from [WeChat Official Platform](https://mp.weixin.qq.com) → Settings & Development → Basic Config.

> **IP 白名单 / IP Whitelist**: 如果启用了 IP 白名单，需将本机出口 IP 加入白名单，否则 API 调用会返回 40125 错误。
> If IP whitelist is enabled, add your machine's outbound IP to the whitelist, otherwise API calls will return error 40125.

### 3. 使用 Usage

```bash
# 查看草稿列表 / List drafts
python wechat_push.py list

# 创建新草稿 / Create a new draft
python wechat_push.py create article.html

# 更新已有草稿 / Update an existing draft
python wechat_push.py update <media_id> article.html

# 删除草稿 / Delete a draft
python wechat_push.py delete <media_id>

# 生成封面图预览 / Generate cover image preview
python wechat_push.py cover "文章标题"
```

---

## 依赖 Dependencies

- Python 3.8+
- `requests`
- `Pillow`

---

## 问题反馈 / Feedback

- 🌐 GitHub：[https://github.com/andy8663/wechat-oa](https://github.com/andy8663/wechat-oa)
- 📧 邮箱：`andy8663@126.com`
