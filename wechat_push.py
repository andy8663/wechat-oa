#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
微信公众号草稿推送工具 v2.0
支持：创建草稿、更新草稿、查看草稿列表、删除草稿
自动生成封面图（2.35:1 比例）
"""
import sys
import io
import os

# Windows 下：stdout 用 UTF-8 编码。
# PowerShell 通过 OpenClaw exec 调用时，OpenClaw 按 UTF-8 解码；
# 直接在 cmd/PowerShell 里运行时，Windows Terminal 支持 UTF-8。
# 如果 reconfigure 失败（老版本 Python），用兼容方案。
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        # Python 3.6 之前版本，fallback
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import json
import requests
import re
import os
from pathlib import Path
from datetime import datetime

# 尝试导入 PIL，用于生成封面图
try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# 配置 - 从 config.json 读取
CONFIG_FILE = Path(__file__).parent / "config.json"

def load_config():
    """从 config.json 加载配置"""
    default_config = {
        "APP_ID": "",
        "APP_SECRET": "",
        "author": "Woody"
    }
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                user_config = json.load(f)
            return {**default_config, **user_config}
        except Exception:
            return default_config
    return default_config

CONFIG = load_config()
APP_ID = CONFIG.get("APP_ID", "")
APP_SECRET = CONFIG.get("APP_SECRET", "")
DEFAULT_AUTHOR = CONFIG.get("author", "Woody")

API_TOKEN_URL = "https://api.weixin.qq.com/cgi-bin/token"
API_DRAFT_ADD = "https://api.weixin.qq.com/cgi-bin/draft/add"
API_DRAFT_GET = "https://api.weixin.qq.com/cgi-bin/draft/get"
API_DRAFT_UPDATE = "https://api.weixin.qq.com/cgi-bin/draft/update"
API_DRAFT_DELETE = "https://api.weixin.qq.com/cgi-bin/draft/delete"
API_DRAFT_BATCHGET = "https://api.weixin.qq.com/cgi-bin/draft/batchget"
API_MATERIAL_ADD = "https://api.weixin.qq.com/cgi-bin/material/add_material"
API_MATERIAL_GET = "https://api.weixin.qq.com/cgi-bin/material/get_material"
API_MATERIAL_DEL = "https://api.weixin.qq.com/cgi-bin/material/del_material"
API_MATERIAL_COUNT = "https://api.weixin.qq.com/cgi-bin/material/get_materialcount"
API_MATERIAL_BATCHGET = "https://api.weixin.qq.com/cgi-bin/material/batchget_material"
API_PUBLISHED_BATCHGET = "https://api.weixin.qq.com/cgi-bin/material/batchget_material"


def hex_to_rgb(hex_color):
    """十六进制颜色转RGB元组"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def generate_cover(title: str, output_path: str) -> str:
    """
    根据文章标题生成封面图
    比例: 2.35:1 (900x383)
    """
    if not HAS_PIL:
        raise Exception("需要安装 Pillow: pip install Pillow")

    WIDTH = 900
    HEIGHT = 383

    # 创建图片
    img = Image.new('RGB', (WIDTH, HEIGHT), hex_to_rgb('#1a1a2e'))
    draw = ImageDraw.Draw(img)

    # 渐变背景
    for y in range(HEIGHT):
        ratio = y / HEIGHT
        r1, g1, b1 = hex_to_rgb('#1a1a2e')
        r2, g2, b2 = hex_to_rgb('#16213e')
        r = int(r1 + (r2 - r1) * ratio)
        g = int(g1 + (g2 - g1) * ratio)
        b = int(b1 + (b2 - b1) * ratio)
        draw.line([(0, y), (WIDTH, y)], fill=(r, g, b))

    # 加载字体
    try:
        font_title = ImageFont.truetype('C:/Windows/Fonts/msyhbd.ttc', 34)
        font_subtitle = ImageFont.truetype('C:/Windows/Fonts/msyh.ttc', 18)
        font_small = ImageFont.truetype('C:/Windows/Fonts/msyh.ttc', 11)
    except:
        font_title = ImageFont.load_default()
        font_subtitle = font_title
        font_small = font_title

    # 主标题（截断显示）
    display_title = title
    if len(display_title) > 18:
        display_title = display_title[:18] + "..."

    draw.text((WIDTH//2, 55), display_title, fill=hex_to_rgb('#ffffff'), font=font_title, anchor='mm')

    # 副标题
    draw.text((WIDTH//2, 98), 'AI Agent Skill', fill=hex_to_rgb('#909090'), font=font_subtitle, anchor='mm')

    # 装饰元素
    draw.rectangle([(0, HEIGHT-40), (WIDTH, HEIGHT)], fill='#0a0a1a')
    draw.text((WIDTH//2, HEIGHT-20), 'OpenClaw', fill=hex_to_rgb('#505050'), font=font_small, anchor='mm')
    draw.rectangle([(0, 0), (WIDTH, 3)], fill=hex_to_rgb('#1a73e8'))

    # 确保目录存在
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    img.save(output_path, 'PNG')
    return output_path


def get_access_token():
    """获取 access_token"""
    if not APP_ID or not APP_SECRET:
        raise Exception("未配置凭证！请将 config.example.json 复制为 config.json，并填入 AppID 和 AppSecret。")
    resp = requests.get(API_TOKEN_URL, params={
        "grant_type": "client_credential",
        "appid": APP_ID,
        "secret": APP_SECRET
    }, timeout=30)
    data = resp.json()
    if "access_token" in data:
        return data["access_token"]
    else:
        raise Exception(f"获取token失败: {data}")


def parse_html_article(html_path):
    """解析 HTML 文件，提取标题和正文"""
    with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    # 提取标题
    title_match = re.search(r'<title>(.*?)</title>', content, re.IGNORECASE)
    title = title_match.group(1).strip() if title_match else "无标题"

    # 微信公众号标题限制：最多64字节（约32个中文字符）
    b = title.encode('utf-8')
    if len(b) > 64:
        title = b[:64].decode('utf-8', errors='ignore')

    # 提取正文（移除标题、样式、脚本）
    body_match = re.search(r'<body[^>]*>(.*?)</body>', content, re.IGNORECASE | re.DOTALL)
    if body_match:
        body = body_match.group(1)
    else:
        # 没有 <body> 时：取 <html> 内容并去掉 <head>
        html_match = re.search(r'<html[^>]*>(.*?)</html>', content, re.IGNORECASE | re.DOTALL)
        if html_match:
            body = html_match.group(1)
        else:
            # 纯片段：去掉所有 <title> 和 <head>
            body = content
        body = re.sub(r'<head[^>]*>.*?</head>', '', body, flags=re.IGNORECASE | re.DOTALL)
        body = re.sub(r'<title[^>]*>.*?</title>', '', body, flags=re.IGNORECASE | re.DOTALL)

    # 移除 script（style 标签保留，供 css_to_inline 转换为行内样式）
    body = re.sub(r'<script[^>]*>.*?</script>', '', body, flags=re.IGNORECASE | re.DOTALL)
    # body = re.sub(r'<style[^>]*>.*?</style>', '', body, ...)  # 注释：保留，css_to_inline 会转换

    # 修复列表项之间的多余空行：</li> 和 <li> 之间不能有换行/空格
    # 同时处理 </li> 与 </ul>/</ol> 之间的空行
    body = re.sub(r'</li>\s*\n\s*<li', '</li><li', body, flags=re.IGNORECASE)
    body = re.sub(r'</li>\s*\n\s*</ul>', '</li></ul>', body, flags=re.IGNORECASE)
    body = re.sub(r'</li>\s*\n\s*</ol>', '</li></ol>', body, flags=re.IGNORECASE)
    body = re.sub(r'<ul[^>]*>\s*\n\s*<li', lambda m: m.group(0).replace('\n', '').replace('  ', ''), body, flags=re.IGNORECASE)
    body = re.sub(r'<ol[^>]*>\s*\n\s*<li', lambda m: m.group(0).replace('\n', '').replace('  ', ''), body, flags=re.IGNORECASE)

    # 清理所有多余空白：多个空格/换行合并为一个
    body = re.sub(r'[ \t]+', ' ', body)       # 多余空格
    body = re.sub(r'\n{2,}', '\n', body)        # 连续换行合并
    body = re.sub(r' *\n *', '\n', body)       # 每行首尾空格
    body = re.sub(r'>\n+<', '><', body)        # 标签间多余空行

    return title.strip(), body.strip()


def upload_image(access_token, image_path):
    """上传永久图片素材"""
    if not Path(image_path).exists():
        raise FileNotFoundError(f"图片文件不存在: {image_path}")
    url = f"{API_MATERIAL_ADD}?access_token={access_token}&type=image"
    with open(image_path, 'rb') as f:
        files = {'media': (Path(image_path).name, f, 'image/png')}
        resp = requests.post(url, files=files, timeout=60)
    data = resp.json()
    if "media_id" in data:
        return data["media_id"]
    else:
        raise Exception(f"上传图片失败: {data}")


def generate_cover_and_upload(access_token, title, html_path):
    """
    生成封面图并上传，返回 thumb_media_id。

    封面图保存到 skill 目录（而非 HTML 所在目录），
    避免 Windows 中文路径导致 PIL img.save() 失败。
    如果封面图生成失败，返回空字符串；草稿仍可创建（需手动补封面）。
    """
    # 固定保存到 skill 目录，文件名不含中文，避免路径编码问题
    skill_dir = Path(__file__).parent
    cover_path = skill_dir / "cover_latest.png"

    print(f"[COVER] 正在生成封面图...")
    cover_ok = False
    try:
        generate_cover(title, str(cover_path))
        print(f"[OK] 封面图已保存: {cover_path}")
        cover_ok = True
    except Exception as e:
        print(f"[WARN] 封面图生成失败: {e}，将跳过封面（需手动补封面图）")

    thumb_media_id = ""
    if cover_ok and cover_path.exists():
        print("[IMG] 正在上传封面图...")
        try:
            thumb_media_id = upload_image(access_token, str(cover_path))
            print(f"[OK] 封面图已上传: {thumb_media_id}")
        except Exception as e:
            print(f"[WARN] 封面上传失败: {e}，将跳过封面")
    else:
        print("[WARN] 封面图文件不存在，跳过上传")

    return thumb_media_id


def css_to_inline(html):
    """
    将 HTML 中 <style> 标签内的 CSS 规则转换为行内样式。
    微信会过滤 <style> 标签，此函数确保样式在行内生效。
    支持：标签选择器、class 选择器、tag.class 组合。
    """
    css_map = {}
    html = re.sub(
        r'<style[^>]*>(.*?)</style>',
        lambda m2: _parse_css_rules(m2.group(1), css_map),
        html,
        flags=re.DOTALL
    )

    if not css_map:
        return html

    # 匹配标签，处理属性值内含 > 的情况
    def tag_replacer(m):
        raw = m.group(0)
        # 跳过闭合标签
        if raw.startswith('</'):
            return raw
        tag_m = re.match(r'</?([a-zA-Z][a-zA-Z0-9-]*)', raw)
        if not tag_m:
            return raw
        tag = tag_m.group(1).lower()
        after = raw[tag_m.end():]
        # 找第一个不在引号内的 >
        i = 0
        in_quote = False
        quote_char = None
        while i < len(after):
            c = after[i]
            if not in_quote and c in ('"', "'"):
                in_quote = True
                quote_char = c
            elif in_quote and c == quote_char:
                in_quote = False
                quote_char = None
            elif not in_quote and c == '>':
                break
            i += 1
        attrs_str = after[:i]
        trailing_slash = ''
        if attrs_str.rstrip().endswith('/'):
            trailing_slash = '/'
            attrs_str = attrs_str.rstrip()[:-1]
        attrs = attrs_str

        cls_match = re.search(r'class=["\']([^"\']+)["\']', attrs)
        cls = cls_match.group(1).split()[0] if cls_match else ""

        applied = []
        key = f"{tag}.{cls}" if cls else None
        if key and key in css_map:
            applied.append(css_map[key])
        if cls and f".{cls}" in css_map:
            applied.append(css_map[f".{cls}"])
        if tag in css_map:
            applied.append(css_map[tag])

        existing_style = re.search(r'style=["\']([^"\']*)["\']', attrs)
        if existing_style:
            applied.insert(0, existing_style.group(1))

        new_style = "; ".join(applied).rstrip("; ")
        if new_style:
            attrs = re.sub(r' style=["\'][^"\']*["\']', '', attrs)
            attrs = attrs + f' style="{new_style}"'

        slash = '/' if raw.startswith('</') else ''
        return f"<{slash}{tag}{attrs}{trailing_slash}>"

    html = re.sub(r'</?[a-zA-Z][a-zA-Z0-9-]*(?:\s+[^>]*)?/?>', tag_replacer, html)
    return html
def _parse_css_rules(css_text, css_map):
    """解析 CSS 文本，将规则存入 css_map。返回空字符串（清除 style 标签）。"""
    for block in css_text.split('}'):
        block = block.strip()
        if not block or '{' not in block:
            continue
        selector, _, props = block.partition('{')
        selector = selector.strip()
        props = props.strip().rstrip(';').strip()
        if not selector or not props:
            continue
        if ':' in selector.split()[0] or '[' in selector or '#' in selector:
            continue
        css_map[selector] = props
    return ""




def _extract_digest(content, max_len=80):
    """
    Smart digest extraction from HTML content.
    Strategy:
      1. Strip HTML tags and entities, split into lines
      2. Skip headings, list items, numbered lines, very short lines
      3. Skip lines that are the same as the title (avoid title duplication)
      4. Boost lines with functional keywords
      5. Pick the best scored paragraph and truncate at natural break points
    """
    # Strip HTML tags
    plain = re.sub(r'<[^>]+>', '', content)
    # Strip HTML entities (e.g. &ldquo; &rdquo; &mdash;)
    plain = re.sub(r'&[a-zA-Z]+;', ' ', plain)
    plain = re.sub(r'\s{2,}', '\n', plain)
    lines = [l.strip() for l in plain.split('\n') if l.strip()]

    # Functional keywords that indicate substantive content
    boost_kws = [
        '\u652f\u6301', '\u5e2e\u4f60', '\u80fd\u591f', '\u53ef\u4ee5', '\u53ea\u9700',
        '\u4e00\u952e', '\u81ea\u52a8', '\u89e3\u51b3', '\u544a\u522b', '\u4ece\u6b64',
        '\u5feb\u901f', '\u8f7b\u677e', '\u5b9e\u73b0', '\u641e\u5b9a', '\u514d\u8d39',
        'skill', 'openclaw', '\u63a8\u9001', '\u8349\u7a3f',
    ]

    best = ''
    best_score = -1

    for line in lines:
        if len(line) < 15:
            continue
        # Skip numbered headings like "01", "1."
        if re.match(r'^\d+[\.\u3001\s]', line):
            continue
        # Skip lines with high quote/bracket ratio
        punct = '\'"\"\u300c\u300d\u300e\u300f'
        punct_count = sum(1 for c in line if c in punct)
        if punct_count > len(line) * 0.15:
            continue
        # Score
        score = sum(3 for kw in boost_kws if kw in line)
        if 30 <= len(line) <= 120:
            score += 1
        if score > best_score:
            best_score = score
            best = line

    # Fallback: first meaningful non-heading paragraph
    if not best:
        for line in lines:
            if len(line) >= 20 and not re.match(r'^\d+[\.\u3001\s]', line):
                best = line
                break

    if not best:
        best = plain[:max_len].strip()

    # Truncate, prefer natural break points (comma/period)
    if len(best) > max_len:
        cut = max_len
        for punct in ['\u3002', '\uff0c', '\uff1b', '\uff01', '\uff1f', '. ', ', ']:
            idx = best.rfind(punct, 0, max_len)
            if idx > max_len * 0.6:
                cut = idx + len(punct)
                break
        best = best[:cut].rstrip('\uff0c\u3001')

    return best



def build_article(title, content, thumb_media_id, author=None):
    """构建图文消息数据"""
    if author is None:
        author = CONFIG.get("author", "Woody")

    # 微信过滤 <style> 标签，转为行内样式
    content = css_to_inline(content)

    # 摘要：智能提取
    digest = _extract_digest(content)

    return {
        "title": title,
        "author": author,
        "content": content,
        "content_source_url": "",
        "digest": digest,
        "thumb_media_id": thumb_media_id,
    }


def save_draft_record(title, media_id):
    """保存草稿记录"""
    record_file = Path(__file__).parent / "draft_ids.txt"
    with open(record_file, 'a', encoding='utf-8') as f:
        f.write(f"{datetime.now().isoformat()} | {title} | {media_id}\n")
    print(f"[INFO] 草稿记录已保存到 {record_file}")


# ========== 草稿操作 ==========

def draft_create(html_path, force_cover=False):
    """创建新草稿

    Args:
        html_path: HTML 文件路径
        force_cover: 是否强制重新生成封面，默认 False（新建草稿无旧封面可复用）
    """
    title, content = parse_html_article(html_path)
    print(f"[TITLE] {title}")
    print(f"[LENGTH] {len(content)} chars")

    access_token = get_access_token()

    # 封面：生成并上传到 skill 目录（避免中文路径），封面失败则报错退出
    thumb_media_id = generate_cover_and_upload(access_token, title, html_path)
    if not thumb_media_id:
        raise Exception(
            "[ERROR] 封面图生成/上传失败，无法创建草稿（草稿必须有封面图）。"
            "请确保已安装 Pillow: pip install Pillow"
        )

    article = build_article(title, content, thumb_media_id)

    # ★ 关键修复：必须用 data=json.dumps(..., ensure_ascii=False).encode('utf-8')
    # 绝对不能用 json=payload（默认 ASCII 转义，中文会变成 \\uXXXX 导致乱码）
    json_data = json.dumps({"articles": [article]}, ensure_ascii=False).encode('utf-8')
    url = f"{API_DRAFT_ADD}?access_token={access_token}"
    resp = requests.post(url, data=json_data, headers={'Content-Type': 'application/json; charset=utf-8'}, timeout=30)
    data = resp.json()

    if "media_id" in data:
        media_id = data["media_id"]
        print(f"\n[OK] 草稿创建成功!")
        print(f"[MEDIA_ID] {media_id}")
        save_draft_record(title, media_id)
        return media_id
    else:
        raise Exception(f"创建草稿失败: {data}")




def _get_draft_thumb_media_id(access_token, media_id):
    """查询已有草稿的封面 thumb_media_id（用于 update 时复用封面）"""
    url = API_DRAFT_GET + "?access_token=" + access_token
    json_data = json.dumps({"media_id": media_id}, ensure_ascii=False).encode("utf-8")
    resp = requests.post(url, data=json_data, headers={"Content-Type": "application/json; charset=utf-8"}, timeout=30)
    data = resp.json()
    items = data.get("news_item", [])
    if items:
        return items[0].get("thumb_media_id", "")
    return ""


def draft_update(media_id, html_path, force_cover=False):
    """更新已有草稿

    Args:
        media_id: 草稿 media_id
        html_path: HTML 文件路径
        force_cover: 是否强制重新生成封面，默认 False（复用已有封面）
    """
    title, content = parse_html_article(html_path)
    print(f"[TITLE] {title}")
    print(f"[LENGTH] {len(content)} chars")
    print(f"[TARGET] 更新草稿: {media_id}")

    access_token = get_access_token()

    # 封面逻辑：默认复用已有草稿封面，--force-cover 才重新生成
    if force_cover:
        thumb_media_id = generate_cover_and_upload(access_token, title, html_path)
        print(f"[COVER] 已强制重新生成封面")
    else:
        # 尝试获取已有草稿的封面 media_id
        thumb_media_id = _get_draft_thumb_media_id(access_token, media_id)
        if thumb_media_id:
            print(f"[COVER] 复用已有封面: {thumb_media_id}")
        else:
            # 兜底：没有封面则生成
            thumb_media_id = generate_cover_and_upload(access_token, title, html_path)
            print(f"[COVER] 无已有封面，已生成新封面")

    article = build_article(title, content, thumb_media_id)

    json_data = json.dumps({
        "media_id": media_id,
        "index": 0,
        "articles": article
    }, ensure_ascii=False).encode('utf-8')

    url = f"{API_DRAFT_UPDATE}?access_token={access_token}"
    resp = requests.post(url, data=json_data, headers={'Content-Type': 'application/json; charset=utf-8'}, timeout=30)
    data = resp.json()

    if data.get("errcode") == 0:
        print(f"\n[OK] 草稿更新成功! Media ID: {media_id}")
    else:
        raise Exception(f"更新草稿失败: {data}")


def draft_list(count=10, offset=0):
    """获取草稿列表"""
    access_token = get_access_token()
    url = f"{API_DRAFT_BATCHGET}?access_token={access_token}"

    json_data = json.dumps({
        "offset": offset,
        "count": count,
        "no_content": 0
    }, ensure_ascii=False).encode('utf-8')

    resp = requests.post(url, data=json_data, headers={'Content-Type': 'application/json; charset=utf-8'}, timeout=30)
    data = resp.json()

    errcode = data.get("errcode", 0)
    if errcode != 0:
        raise Exception(f"获取草稿列表失败: {data}")

    items = data.get("item", [])
    total = data.get("total_count", 0)
    print(f"\n[草稿箱] 共 {total} 篇草稿 (显示 {len(items)} 篇):\n")
    for i, item in enumerate(items):
        media_id = item.get("media_id", "N/A")
        articles = item.get("content", {}).get("news_item", [])
        if articles:
            art = articles[0]
            title = art.get("title", "无标题")
            # digest 可能有换行符（HTML 中的空白符），清理掉
            digest_raw = art.get("digest", "")
            digest = re.sub(r'[\r\n]+', ' ', digest_raw).strip()
            update_time = datetime.fromtimestamp(item.get("update_time", 0)).strftime("%Y-%m-%d %H:%M")
            # 输出一行，避免 Windows PowerShell 空行问题
            print(f"  [{i+1}] {title}  |  {update_time}  |  {media_id}")
            if digest:
                print(f"      摘要: {digest[:60]}{'...' if len(digest) > 60 else ''}")
    return items


def draft_delete(media_id):
    """删除草稿"""
    access_token = get_access_token()
    url = f"{API_DRAFT_DELETE}?access_token={access_token}"

    json_data = json.dumps({"media_id": media_id}, ensure_ascii=False).encode('utf-8')
    resp = requests.post(url, data=json_data, headers={'Content-Type': 'application/json; charset=utf-8'}, timeout=30)
    data = resp.json()

    if data.get("errcode") == 0:
        print(f"[OK] 草稿已删除: {media_id}")
    else:
        raise Exception(f"删除草稿失败: {data}")


def published_list(count=10, offset=0):
    """获取已发布文章列表"""
    access_token = get_access_token()
    url = f"{API_PUBLISHED_BATCHGET}?access_token={access_token}"

    json_data = json.dumps({
        "type": "news",
        "offset": offset,
        "count": count
    }, ensure_ascii=False).encode('utf-8')

    resp = requests.post(url, data=json_data, headers={'Content-Type': 'application/json; charset=utf-8'}, timeout=30)
    data = resp.json()

    errcode = data.get("errcode", 0)
    if errcode != 0:
        raise Exception(f"获取已发布列表失败: {data}")

    items = data.get("item", [])
    total = data.get("total_count", 0)
    print(f"\n[已发布文章] 共 {total} 篇 (显示 {len(items)} 篇):\n")
    for i, item in enumerate(items):
        articles = item.get("content", {}).get("news_item", [])
        if articles:
            art = articles[0]
            title = art.get("title", "无标题")
            update_time = datetime.fromtimestamp(item.get("update_time", 0)).strftime("%Y-%m-%d %H:%M")
            print(f"  [{i+1}] {title}  |  {update_time}")
    return items


def material_upload(image_path):
    """上传永久素材（图片）"""
    access_token = get_access_token()

    if not Path(image_path).exists():
        raise FileNotFoundError(f"文件不存在: {image_path}")

    url = f"{API_MATERIAL_ADD}?access_token={access_token}&type=image"
    with open(image_path, 'rb') as f:
        files = {'media': (Path(image_path).name, f, 'image/png')}
        resp = requests.post(url, files=files, timeout=60)
    data = resp.json()

    if "media_id" in data:
        media_id = data["media_id"]
        url_out = data.get("url", "")
        print(f"[OK] 永久素材上传成功!")
        print(f"[MEDIA_ID] {media_id}")
        if url_out:
            print(f"[URL] {url_out}")
        return media_id
    else:
        raise Exception(f"上传素材失败: {data}")


def material_count():
    """获取各类永久素材总数"""
    access_token = get_access_token()
    url = f"{API_MATERIAL_COUNT}?access_token={access_token}"
    resp = errwrap_get(url)
    data = resp.json()

    errcode = data.get("errcode", 0)
    if errcode != 0:
        raise Exception(f"获取素材总数失败: {data}")

    voice_count = data.get("voice_count", 0)
    video_count = data.get("video_count", 0)
    image_count = data.get("image_count", 0)
    news_count = data.get("news_count", 0)
    total = voice_count + video_count + image_count + news_count

    print(f"\n[永久素材统计]")
    print(f"  语音: {voice_count}")
    print(f"  视频: {video_count}")
    print(f"  图片: {image_count}")
    print(f"  图文: {news_count}")
    print(f"  ─────────────────")
    print(f"  合计: {total}")
    return data


def material_list(mtype="image", count=20, offset=0):
    """批量获取永久素材列表

    Args:
        mtype: 素材类型，支持 image / video / voice / news
        count: 每页数量，默认20
        offset: 偏移量，默认0
    """
    access_token = get_access_token()
    url = f"{API_MATERIAL_BATCHGET}?access_token={access_token}"

    json_data = json.dumps({
        "type": mtype,
        "offset": offset,
        "count": count
    }, ensure_ascii=False).encode('utf-8')

    resp = requests.post(url, data=json_data, headers={'Content-Type': 'application/json; charset=utf-8'}, timeout=30)
    data = resp.json()

    errcode = data.get("errcode", 0)
    if errcode != 0:
        raise Exception(f"获取素材列表失败: {data}")

    items = data.get("item", [])
    total = data.get("total_count", 0)

    type_label = {"image": "图片", "video": "视频", "voice": "语音", "news": "图文"}
    label = type_label.get(mtype, mtype)

    print(f"\n[永久素材列表] 类型:{label}  共{total}个 (显示{len(items)}个):\n")
    for i, item in enumerate(items):
        media_id = item.get("media_id", "N/A")
        update_time = datetime.fromtimestamp(item.get("update_time", 0)).strftime("%Y-%m-%d %H:%M")

        if mtype == "image":
            name = item.get("name", "N/A")
            url_out = item.get("url", "")
            print(f"  [{i+1}] {name}  |  {update_time}  |  {media_id}")
            if url_out:
                print(f"      URL: {url_out}")
        elif mtype == "video":
            name = item.get("name", "N/A")
            print(f"  [{i+1}] {name}  |  {update_time}  |  {media_id}")
        elif mtype == "voice":
            name = item.get("name", "N/A")
            print(f"  [{i+1}] {name}  |  {update_time}  |  {media_id}")
        elif mtype == "news":
            articles = item.get("content", {}).get("news_item", [])
            if articles:
                title = articles[0].get("title", "无标题")
                print(f"  [{i+1}] {title}  |  {update_time}  |  {media_id}")
        print()
    return items


def errwrap_get(url):
    """兼容 requests.get 的包装（API 实际返回 JSON）"""
    return requests.get(url, timeout=30)


def print_usage():
    print("""
微信公众号草稿推送工具 v2.0

用法:
  python wechat_push.py list                          查看草稿列表
  python wechat_push.py create <html文件>              创建新草稿
  python wechat_push.py update <media_id> <html文件> [--force-cover]   更新已有草稿（默认复用封面）
  python wechat_push.py delete <media_id>              删除草稿
  python wechat_push.py upload <图片文件>              上传永久素材（图片）
  python wechat_push.py materialcount                  获取永久素材总数统计
  python wechat_push.py materials [type] [count] [offset]  批量获取永久素材列表
                                                          type: image/video/voice/news，默认 image
  python wechat_push.py materialdel <media_id>         删除永久素材
  python wechat_push.py published                      获取已发布文章列表

示例:
  python wechat_push.py list
  python wechat_push.py create article.html
  python wechat_push.py materialcount
  python wechat_push.py materials image 20 0
  python wechat_push.py materials news 10 0
""")


def main():
    args = sys.argv[1:]

    if len(args) == 0 or args[0] in ('-h', '--help', 'help'):
        print_usage()
        sys.exit(0)

    cmd = args[0].lower()

    try:
        if cmd == 'list':
            draft_list()

        elif cmd == 'create':
            if len(args) < 2:
                print("[ERROR] 请指定HTML文件路径")
                print("用法: python wechat_push.py create <html文件> [--force-cover]")
                sys.exit(1)
            html_path = args[1]
            remaining = args[2:]
            force_cover = '--force-cover' in remaining
            draft_create(html_path, force_cover=force_cover)

        elif cmd == 'update':
            if len(args) < 2:
                print("[ERROR] 请指定 media_id 和 HTML文件路径")
                print("用法: python wechat_push.py update <media_id> <html文件> [--force-cover]")
                sys.exit(1)
            media_id = args[1]
            html_path = args[2] if len(args) > 2 else None
            # 解析剩余参数，支持 --force-cover
            remaining = args[3:] if len(args) > 3 else args[2:]
            if html_path is None and remaining:
                html_path = remaining[0]
                remaining = remaining[1:]
            force_cover = '--force-cover' in remaining
            if html_path is None:
                print("[ERROR] 请指定 HTML文件路径")
                sys.exit(1)
            draft_update(media_id, html_path, force_cover=force_cover)

        elif cmd == 'delete':
            if len(args) < 2:
                print("[ERROR] 请指定 media_id")
                print("用法: python wechat_push.py delete <media_id>")
                sys.exit(1)
            draft_delete(args[1])

        elif cmd == 'upload':
            if len(args) < 2:
                print("[ERROR] 请指定图片文件路径")
                print("用法: python wechat_push.py upload <图片文件>")
                sys.exit(1)
            material_upload(args[1])

        elif cmd == 'published':
            published_list()

        elif cmd == 'materialcount':
            material_count()

        elif cmd == 'materials':
            # 解析: materials [type] [count] [offset]
            mtype = "image"
            count = 20
            offset = 0
            if len(args) >= 2:
                mtype = args[1]
            if len(args) >= 3:
                count = int(args[2])
            if len(args) >= 4:
                offset = int(args[3])
            material_list(mtype, count, offset)

        elif cmd == 'materialdel':
            if len(args) < 2:
                print("[ERROR] 请指定 media_id")
                print("用法: python wechat_push.py materialdel <media_id>")
                sys.exit(1)
            access_token = get_access_token()
            url = f"{API_MATERIAL_DEL}?access_token={access_token}"
            json_data = json.dumps({"media_id": args[1]}, ensure_ascii=False).encode('utf-8')
            resp = requests.post(url, data=json_data, headers={'Content-Type': 'application/json; charset=utf-8'}, timeout=30)
            data = resp.json()
            if data.get("errcode") == 0:
                print(f"[OK] 永久素材已删除: {args[1]}")
            else:
                raise Exception(f"删除素材失败: {data}")

        else:
            # 兼容旧用法：直接传html文件路径 = 创建新草稿
            if args[0].endswith('.html'):
                print("[INFO] 检测到HTML文件，默认创建新草稿")
                draft_create(args[0])
            else:
                print(f"[ERROR] 未知命令: {cmd}")
                print_usage()
                sys.exit(1)

    except Exception as e:
        print(f"\n[ERROR] {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
