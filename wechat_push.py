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

# 尝试导入 premailer，用于 CSS 内联转换
try:
    from premailer import Premailer
    HAS_PREMAILER = True
except ImportError:
    HAS_PREMAILER = False

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
API_USER_SUMMARY = "https://api.weixin.qq.com/datacube/getusersummary"
API_USER_CUMULATE = "https://api.weixin.qq.com/datacube/getusercumulate"
API_USER_INFO = "https://api.weixin.qq.com/cgi-bin/user/info"
API_USER_LIST = "https://api.weixin.qq.com/cgi-bin/user/get"


def hex_to_rgb(hex_color: str):
    """将 #RRGGBB 格式的颜色转为 (R, G, B)"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def generate_cover(title: str, output_path: str) -> str:
    """
    根据文章标题生成科技风封面图
    比例: 2.35:1 (900x383)
    """
    if not HAS_PIL:
        raise Exception("需要安装 Pillow: pip install Pillow")

    WIDTH = 900
    HEIGHT = 383

    # 深色科技背景
    img = Image.new('RGB', (WIDTH, HEIGHT), (5, 9, 28))
    draw = ImageDraw.Draw(img)

    # ---------- 1. 渐变底层（从上到下：深蓝→纯黑） ----------
    for y in range(HEIGHT):
        ratio = y / HEIGHT
        r = int(5 + (0 - 5) * ratio)
        g = int(9 + (0 - 9) * ratio)
        b = int(28 + (0 - 28) * ratio)
        draw.line([(0, y), (WIDTH, y)], fill=(r, g, b))

    # ---------- 2. 细网格线（横向 + 纵向） ----------
    grid_color = (20, 40, 80)
    # 横线每15px一条
    for y in range(0, HEIGHT, 15):
        draw.line([(0, y), (WIDTH, y)], fill=grid_color)
    # 竖线每20px一条
    for x in range(0, WIDTH, 20):
        draw.line([(x, 0), (x, HEIGHT)], fill=grid_color)

    # ---------- 3. 底部数据流色带 ----------
    draw.rectangle([(0, 0), (WIDTH, 4)], fill=(0, 200, 255))   # 顶部霓虹蓝线
    draw.rectangle([(0, 0), (4, HEIGHT)], fill=(0, 180, 240))  # 左侧霓虹蓝线
    draw.rectangle([(WIDTH-4, 0), (WIDTH, HEIGHT)], fill=(0, 80, 160))  # 右侧暗蓝线

    # ---------- 4. 底部科技装饰条 ----------
    draw.rectangle([(0, HEIGHT-60), (WIDTH, HEIGHT)], fill=(0, 8, 25))
    # 底部内嵌线
    draw.line([(0, HEIGHT-60), (WIDTH, HEIGHT-60)], fill=(0, 140, 220))

    # ---------- 5. 左侧竖向标签 ----------
    # 画一个带透明度的侧边标签区
    for x in range(0, 8):
        alpha_ratio = 1 - x / 8
        r, g, b = 0, 180, 255
        draw.line([(x, 0), (x, HEIGHT)], fill=(int(r*alpha_ratio), int(g*alpha_ratio), int(b*alpha_ratio)))

    # ---------- 6. 二进制/十六进制数字装饰（左上/右下角） ----------
    try:
        font_hex = ImageFont.truetype('C:/Windows/Fonts/consola.ttf', 9)
    except:
        font_hex = ImageFont.load_default()

    hex_nums = ['0x00FF', '0xA3', '0x7E', '0x1B', '0xF0', '0x3C',
                '10110', '00101', '11100', '01010', '11001', '00111']
    # 左上角
    for i, hx in enumerate(hex_nums[:6]):
        x = 14 + (i % 2) * 48
        y = 14 + (i // 2) * 13
        draw.text((x, y), hx, fill=(0, 90, 160), font=font_hex)
    # 右下角
    for i, hx in enumerate(hex_nums[6:]):
        x = WIDTH - 46 + (i % 2) * 42
        y = HEIGHT - 48 + (i // 2) * 12
        draw.text((x, y), hx, fill=(0, 70, 130), font=font_hex)

    # ---------- 7. 电路连线装饰（半透明线条从左到右） ----------
    circuit_pts = [
        ((8, 70), (60, 90)), ((60, 90), (120, 65)), ((120, 65), (180, 95)),
        ((180, 95), (230, 75)), ((230, 75), (280, 100)),
    ]
    for (x1, y1), (x2, y2) in circuit_pts:
        draw.line([(x1, y1), (x2, y2)], fill=(0, 110, 200), width=1)
        draw.ellipse([(x1-2, y1-2), (x1+2, y1+2)], fill=(0, 180, 255))  # 节点圆点

    circuit_pts2 = [
        ((WIDTH-8, 280), (WIDTH-70, 300)), ((WIDTH-70, 300), (WIDTH-140, 270)),
        ((WIDTH-140, 270), (WIDTH-200, 295)),
    ]
    for (x1, y1), (x2, y2) in circuit_pts2:
        draw.line([(x1, y1), (x2, y2)], fill=(0, 90, 180), width=1)
        draw.ellipse([(x1-2, y1-2), (x1+2, y1+2)], fill=(0, 150, 230))

    # ---------- 8. 中心主文字区域（带发光效果） ----------
    try:
        font_title = ImageFont.truetype('C:/Windows/Fonts/msyhbd.ttc', 32)
        font_subtitle = ImageFont.truetype('C:/Windows/Fonts/msyh.ttc', 16)
        font_label = ImageFont.truetype('C:/Windows/Fonts/arial.ttf', 11)
    except Exception:
        font_title = ImageFont.load_default()
        font_subtitle = font_title
        font_label = font_title

    # 标题处理（最多28字显示）
    display_title = title
    if len(display_title) > 28:
        display_title = display_title[:28] + '...'

    # 文字区域背景（半透明矩形）
    draw.rectangle([(80, 120), (WIDTH-80, HEIGHT-70)], fill=(5, 9, 28))
    draw.rectangle([(80, 120), (WIDTH-80, 122)], fill=(0, 190, 255))  # 顶边亮线
    draw.rectangle([(80, HEIGHT-70), (WIDTH-80, HEIGHT-68)], fill=(0, 100, 200))  # 底边暗线

    # 主标题（白色）
    draw.text((WIDTH//2, 158), display_title, fill=(255, 255, 255), font=font_title, anchor='mm')

    # 分隔线
    draw.line([(WIDTH//2-80, 195), (WIDTH//2+80, 195)], fill=(0, 170, 240), width=1)

    # 副标题（青色）
    draw.text((WIDTH//2, 215), 'A I   A g e n t   S k i l l', fill=(0, 210, 255), font=font_subtitle, anchor='mm')

    # 底部标签
    draw.text((WIDTH//2, HEIGHT-30), '[  OpenClaw · AI Agent  ]', fill=(0, 120, 200), font=font_label, anchor='mm')

    # ---------- 9. 右上角装饰：信号强度图标 ----------
    bars = [(WIDTH-30, 20), (WIDTH-40, 25), (WIDTH-50, 30), (WIDTH-60, 35)]
    heights = [6, 12, 18, 24]
    for i, (bx, _) in enumerate(bars):
        draw.rectangle([(bx, 38 - heights[i]), (bx+7, 38)], fill=(0, 160+i*15, 255))

    # ---------- 10. 左下角装饰：处理器芯片图标 ----------
    chip_x, chip_y = 20, HEIGHT - 50
    draw.rectangle([(chip_x, chip_y), (chip_x+28, chip_y+28)], fill=(0, 20, 60), outline=(0, 140, 255), width=1)
    draw.rectangle([(chip_x+10, chip_y+10), (chip_x+18, chip_y+18)], fill=(0, 60, 140))
    # 引脚
    for i in range(5):
        draw.line([(chip_x-5, chip_y+2+i*6), (chip_x, chip_y+2+i*6)], fill=(0, 120, 220), width=1)
        draw.line([(chip_x+28, chip_y+2+i*6), (chip_x+33, chip_y+2+i*6)], fill=(0, 120, 220), width=1)
        draw.line([(chip_x+2+i*6, chip_y-5), (chip_x+2+i*6, chip_y)], fill=(0, 120, 220), width=1)
        draw.line([(chip_x+2+i*6, chip_y+28), (chip_x+2+i*6, chip_y+33)], fill=(0, 120, 220), width=1)

    # ---------- 保存 ----------
    dir_path = os.path.dirname(output_path)
    if dir_path:
        os.makedirs(dir_path, exist_ok=True)
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
    """解析 HTML 文件，提取标题、正文和样式"""
    with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    # 提取标题：优先 <title>，其次 <h1>，最后找第一个 <hN> 标题
    title = "无标题"
    for pattern in [
        r'<title[^>]*>(.*?)</title>',
        r'<h1[^>]*>(.*?)</h1>',
        r'<h2[^>]*>(.*?)</h2>',
    ]:
        m = re.search(pattern, content, re.IGNORECASE | re.DOTALL)
        if m:
            candidate = m.group(1).strip()
            # 去掉标签残片
            candidate = re.sub(r'<[^>]+>', '', candidate).strip()
            if candidate and candidate != "无标题":
                title = candidate
                break

    # 微信公众号标题限制：最多64个字符（实测上限，65字符会报 title size out of limit）
    # 原文：最多64字节（约32个中文字符）——错误，微信按字符计，非字节
    if len(title) > 64:
        title = title[:64]

    # 提取 <style> 标签内容（用于 CSS 内联转换）
    style_blocks = re.findall(r'<style[^>]*>(.*?)</style>', content, re.IGNORECASE | re.DOTALL)
    style_content = '\n'.join(style_blocks)

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

    # 移除 script（style 标签已在上面提取，这里从 body 中移除）
    body = re.sub(r'<script[^>]*>.*?</script>', '', body, flags=re.IGNORECASE | re.DOTALL)
    body = re.sub(r'<style[^>]*>.*?</style>', '', body, flags=re.IGNORECASE | re.DOTALL)

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

    return title.strip(), body.strip(), style_content.strip()


def parse_md_article(md_path):
    """
    解析 Markdown 文件，转换为带内联样式的 HTML。

    支持的 Markdown 元素：
      # / ## / ###       标题（h1 / h2 / h3）
      **bold** / *italic* / `code`  行内格式
      > blockquote         引用块
      ---                  分隔线
      - / * 无序列表       ul/li
      1. 有序列表          ol/li
      [text](url)          链接
      ![alt](url)          图片（已脱敏，不生成 img 标签）

    样式严格遵循 design.md 规范：clamp() 响应式字号、677px 容器宽度。
    返回：(title, body, style_content)
    """
    with open(md_path, 'r', encoding='utf-8', errors='ignore') as f:
        md = f.read()

    # ── 1. 提取标题 ──────────────────────────────────────────────────────────
    title = "无标题"
    for m in re.finditer(r'^(#{1,3})\s+(.+)$', md, re.MULTILINE):
        raw = m.group(2).strip()
        raw = re.sub(r'\*\*(.+?)\*\*', r'\1', raw)  # 去掉粗体标记
        raw = re.sub(r'\*(.+?)\*', r'\1', raw)        # 去掉斜体标记
        raw = re.sub(r'`(.+?)`', r'\1', raw)          # 去掉代码标记
        if raw:
            title = raw
            break

    if len(title) > 64:
        title = title[:64]

    # ── 2. 预处理：代码块保护 ────────────────────────────────────────────────
    # 用占位符保护 fenced code block 内容，防止转换逻辑误伤
    code_blocks = []
    def protect_code(m):
        placeholder = f'\x00CODEBLOCK{len(code_blocks)}\x00'
        code_blocks.append(m.group(0))
        return placeholder

    md = re.sub(r'```[\s\S]*?```', protect_code, md)       # fenced code block
    md = re.sub(r'`([^`]+)`', protect_code, md)             # inline code

    # ── 3. 预处理：表格行保护 ────────────────────────────────────────────────
    table_rows = []
    def protect_table(m):
        placeholder = f'\x00TABLEROW{len(table_rows)}\x00'
        table_rows.append(m.group(0))
        return placeholder

    md = re.sub(r'\|.+\|(\n\|[|:\- ]+\|)?', protect_table, md)

    # ── 4. 转行：块级元素 ───────────────────────────────────────────────────
    lines = md.split('\n')
    html_lines = []
    in_ul = False
    in_ol = False

    def _close_ul():
        nonlocal in_ul
        if in_ul:
            html_lines.append('</ul>')
            in_ul = False

    def _close_ol():
        nonlocal in_ol
        if in_ol:
            html_lines.append('</ol>')
            in_ol = False

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # 空行
        if not stripped:
            _close_ul()
            _close_ol()
            i += 1
            continue

        # 分隔线
        if re.match(r'^(-{3,}|\*{3,}|_{3,})$', stripped):
            _close_ul()
            _close_ol()
            html_lines.append('<hr>')
            i += 1
            continue

        # 标题
        hm = re.match(r'^(#{1,3})\s+(.+)$', stripped)
        if hm:
            _close_ul()
            _close_ol()
            level = len(hm.group(1))
            inner = hm.group(2)
            inner = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', inner)
            inner = re.sub(r'\*(.+?)\*', r'<em>\1</em>', inner)
            inner = re.sub(r'`([^`]+)`', r'<code>\1</code>', inner)
            html_lines.append(f'<h{level}>{inner}</h{level}>')
            i += 1
            continue

        # 无序列表项（- 或 * 开头）
        li_m = re.match(r'^([-*+])\s+(.+)$', stripped)
        if li_m:
            if not in_ul:
                _close_ol()
                html_lines.append('<ul>')
                in_ul = True
            item = li_m.group(2)
            item = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', item)
            item = re.sub(r'\*(.+?)\*', r'<em>\1</em>', item)
            item = re.sub(r'`([^`]+)`', r'<code>\1</code>', item)
            html_lines.append(f'<li>{item}</li>')
            i += 1
            continue

        # 有序列表项（1. 开头）
        ol_m = re.match(r'^\d+\.\s+(.+)$', stripped)
        if ol_m:
            if not in_ol:
                _close_ul()
                html_lines.append('<ol>')
                in_ol = True
            item = ol_m.group(1)
            item = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', item)
            item = re.sub(r'\*(.+?)\*', r'<em>\1</em>', item)
            item = re.sub(r'`([^`]+)`', r'<code>\1</code>', item)
            html_lines.append(f'<li>{item}</li>')
            i += 1
            continue

        # 引用块
        if stripped.startswith('>'):
            _close_ul()
            _close_ol()
            quote_content = stripped.lstrip('>').strip()
            quote_content = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', quote_content)
            quote_content = re.sub(r'\*(.+?)\*', r'<em>\1</em>', quote_content)
            quote_content = re.sub(r'`([^`]+)`', r'<code>\1</code>', quote_content)
            html_lines.append(f'<blockquote><p>{quote_content}</p></blockquote>')
            i += 1
            continue

        # 普通段落
        _close_ul()
        _close_ol()
        para = stripped
        para = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', para)
        para = re.sub(r'\*(.+?)\*', r'<em>\1</em>', para)
        para = re.sub(r'`([^`]+)`', r'<code>\1</code>', para)
        para = re.sub(r'\[([^\]]+)\]\(([^\)]+)\)', r'<a href="\2">\1</a>', para)
        para = re.sub(r'!\[([^\]]*)\]\([^\)]+\)', '', para)  # 去掉图片 markdown
        html_lines.append(f'<p>{para}</p>')

        i += 1

    # 关闭未关闭的列表
    _close_ul()
    _close_ol()

    body = '\n'.join(html_lines)

    # ── 5. 恢复保护的内容 ───────────────────────────────────────────────────
    for idx, block in enumerate(code_blocks):
        body = body.replace(f'\x00CODEBLOCK{idx}\x00', block)

    for idx, row in enumerate(table_rows):
        # 简单表格还原为 HTML table
        body = body.replace(f'\x00TABLEROW{idx}\x00', _render_table_row(row))

    # ── 6. 组装 style（遵循 design.md） ─────────────────────────────────────
    style_content = """
.design-container { width: 677px; max-width: 100%; margin: 0 auto; box-sizing: border-box; background-color: #fff; }
.content-container { padding: 0; margin: 0; }
h1 { font-size: clamp(18px, 2vw, 20px); font-weight: bold; text-align: center; margin: 20px 0 14px; }
h2 { font-size: clamp(17px, 1.8vw, 18px); font-weight: bold; margin: 18px 0 12px; }
h3 { font-size: clamp(15px, 1.6vw, 16px); font-weight: bold; margin: 14px 0 10px; }
p  { font-size: clamp(15px, 1.4vw, 16px); line-height: 1.8; margin: 8px 0; }
strong { font-weight: bold; }
em { font-style: italic; }
code { background-color: #f5f5f5; padding: 2px 6px; border-radius: 4px; font-family: Consolas, monospace; font-size: 0.9em; }
blockquote { border-left: 4px solid #3498db; background-color: #f0f7ff; padding: 10px 16px; margin: 10px 0; }
blockquote p { color: #555; margin: 0; font-size: clamp(14px, 1.3vw, 15px); }
ul, ol { padding-left: 24px; margin: 8px 0; }
li { font-size: clamp(15px, 1.4vw, 16px); line-height: 1.8; margin: 4px 0; }
a { color: #3498db; text-decoration: none; }
hr { border: none; border-top: 1px solid #e0e0e0; margin: 20px 0; }
table { width: 100%; border-collapse: collapse; font-size: clamp(13px, 1.1vw, 14px); margin: 10px 0; }
th { background-color: #d6eaf8; padding: 8px; text-align: left; }
td { padding: 8px; border-bottom: 1px solid #eee; }
""".strip()

    return title.strip(), body.strip(), style_content


def _render_table_row(md_row):
    """将 Markdown 表格行（含分隔行）转为 HTML <table>"""
    cells = [c.strip() for c in md_row.strip('|').split('|')]
    if not cells:
        return ''

    # 判断是否为分隔行（只含 - : |）
    if all(re.match(r'^[:\- ]+$', c) for c in cells):
        return ''

    is_header = not any(re.match(r'^[:\- ]+$', c) for c in cells)
    tag = 'th' if is_header else 'td'
    cols = ''.join(f'<{tag}>{c}</{tag}>' for c in cells)
    return f'<tr>{cols}</tr>'


def parse_file(file_path):
    """
    通用文件解析入口，自动识别后缀调用对应解析器。

    支持：.html / .htm → HTML 解析器
          .md / .markdown → Markdown 解析器

    返回：(title, body, style_content)
    """
    suffix = Path(file_path).suffix.lower()
    if suffix in ('.html', '.htm'):
        return parse_html_article(file_path)
    elif suffix in ('.md', '.markdown'):
        return parse_md_article(file_path)
    else:
        raise ValueError(f"不支持的文件格式: {suffix}（仅支持 .html 和 .md）")


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

    封面图统一保存到 skill 目录下的 TMP/ 子目录（已加入 .gitignore），
    避免中文路径问题，不污染 Git。
    如果封面图生成失败，返回空字符串；草稿仍可创建（需手动补封面）。
    """
    skill_dir = Path(__file__).parent
    tmp_dir = skill_dir / "TMP"
    tmp_dir.mkdir(exist_ok=True)
    # 文件名不含中文，避免路径编码问题
    safe_name = re.sub(r'[^\w\u4e00-\u9fff]', '_', title)[:20]
    cover_path = tmp_dir / f"cover_{safe_name}.png"

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
    
    优先使用 premailer 库（如果已安装），否则使用原生实现。
    """
    if HAS_PREMAILER:
        try:
            # 使用 premailer 进行专业的 CSS 内联转换
            p = Premailer(
                html,
                remove_classes=False,  # 保留 class 属性
                strip_important=False,  # 保留 !important
                include_star_selectors=False,  # 忽略 * 选择器
                disable_link_rewrites=True  # 禁用自动链接重写，保留原始链接
            )
            return p.transform()
        except Exception as e:
            print(f"[WARN] premailer 转换失败，回退到原生实现: {e}")
            return _css_to_inline_native(html)
    else:
        return _css_to_inline_native(html)

def _css_to_inline_native(html):
    """
    原生 CSS 内联转换实现（作为 premailer 的 fallback）。
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



def build_article(title, content, thumb_media_id, style_content="", author=None):
    """构建图文消息数据"""
    if author is None:
        author = CONFIG.get("author", "Woody")

    # 微信过滤 <style> 标签，转为行内样式
    # 将 style 和 body 组合成完整 HTML 供 premailer 处理
    if style_content:
        full_html = f'<html><head><style>{style_content}</style></head><body>{content}</body></html>'
    else:
        full_html = content
    content = css_to_inline(full_html)
    # 如果返回的是完整 HTML，提取 body 内容
    body_match = re.search(r'<body[^>]*>(.*?)</body>', content, re.IGNORECASE | re.DOTALL)
    if body_match:
        content = body_match.group(1)

    # 清理 premailer 产生的标签间换行和多余空格
    # 修复列表项之间的空白
    content = re.sub(r'<(li|ul|ol|p|div|h[1-6]|section|article|header|footer|nav|aside|main|figure|figcaption|blockquote|pre|code|table|thead|tbody|tr|td|th|caption)\s*>\s*\n\s*<', r'<\1><', content, flags=re.IGNORECASE)
    content = re.sub(r'>\s*\n\s*<', '><', content)  # 全局标签间换行
    content = re.sub(r'\n\s*<', '<', content)  # 行首标签前的换行
    content = re.sub(r'>\s*\n', '>', content)  # 行尾标签后的换行
    content = re.sub(r'[ \t]{2,}', ' ', content)  # 多个空格合并

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
    title, content, style_content = parse_file(html_path)
    print(f"[TITLE] {title}")
    print(f"[LENGTH] {len(content)} chars")
    if style_content:
        print(f"[STYLE] 提取到 {len(style_content)} 字符 CSS 样式")

    access_token = get_access_token()

    # 封面：生成并上传到 skill 目录（避免中文路径），封面失败则报错退出
    thumb_media_id = generate_cover_and_upload(access_token, title, html_path)
    if not thumb_media_id:
        raise Exception(
            "[ERROR] 封面图生成/上传失败，无法创建草稿（草稿必须有封面图）。"
            "请确保已安装 Pillow: pip install Pillow"
        )

    article = build_article(title, content, thumb_media_id, style_content)

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
    title, content, style_content = parse_file(html_path)
    print(f"[TITLE] {title}")
    print(f"[LENGTH] {len(content)} chars")
    if style_content:
        print(f"[STYLE] 提取到 {len(style_content)} 字符 CSS 样式")
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

    article = build_article(title, content, thumb_media_id, style_content)

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


def draft_find(keyword):
    """
    按标题关键词搜索草稿

    Args:
        keyword: 搜索关键词（不区分大小写）
    """
    access_token = get_access_token()
    url = f"{API_DRAFT_BATCHGET}?access_token={access_token}"

    # 每次拉20篇，循环直到拉完
    offset = 0
    page_size = 20
    all_items = []
    while True:
        json_data = json.dumps({
            "offset": offset,
            "count": page_size,
            "no_content": 0
        }, ensure_ascii=False).encode('utf-8')

        resp = requests.post(url, data=json_data,
            headers={'Content-Type': 'application/json; charset=utf-8'}, timeout=30)
        data = resp.json()
        errcode = data.get("errcode", 0)
        if errcode != 0:
            raise Exception(f"搜索草稿失败: {data}")

        items = data.get("item", [])
        all_items.extend(items)
        total = data.get("total_count", 0)

        if len(all_items) >= total or len(items) < page_size:
            break
        offset += page_size

    kw_lower = keyword.lower()
    matched = []
    for item in all_items:
        media_id = item.get("media_id", "")
        articles = item.get("content", {}).get("news_item", [])
        if not articles:
            continue
        art = articles[0]
        title = art.get("title", "")
        update_time = datetime.fromtimestamp(item.get("update_time", 0)).strftime("%Y-%m-%d %H:%M")
        if kw_lower in title.lower():
            matched.append((title, update_time, media_id))

    if not matched:
        print(f"[草稿搜索] 关键词「{keyword}」未找到匹配草稿（共扫描 {len(all_items)} 篇）")
    else:
        print(f"[草稿搜索] 关键词「{keyword}」，共找到 {len(matched)} 篇（共扫描 {len(all_items)} 篇）:\n")
        for i, (title, update_time, media_id) in enumerate(matched):
            print(f"  [{i+1}] {title}")
            print(f"      {update_time}  |  {media_id}")
    return matched


def draft_batch_del(media_ids):
    """
    批量删除草稿

    Args:
        media_ids: media_id 列表
    """
    if not media_ids:
        print("[ERROR] 未指定要删除的草稿 ID")
        return

    print(f"[批量删除] 共 {len(media_ids)} 篇草稿...\n")
    success = 0
    failed = []
    for mid in media_ids:
        try:
            draft_delete(mid)
            success += 1
        except Exception as e:
            failed.append((mid, str(e)))
            print(f"[WARN] 删除失败 {mid}: {e}")

    print(f"\n[结果] 成功 {success} 篇，失败 {len(failed)} 篇")
    if failed:
        print("失败列表:")
        for mid, err in failed:
            print(f"  {mid}: {err}")


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


def user_summary(begin_date, end_date):
    """获取用户增减数据（每日明细）

    Args:
        begin_date: 开始日期，格式 YYYY-MM-DD
        end_date: 结束日期，格式 YYYY-MMDD
    """
    access_token = get_access_token()
    url = f"{API_USER_SUMMARY}?access_token={access_token}"

    json_data = json.dumps({
        "begin_date": begin_date,
        "end_date": end_date
    }, ensure_ascii=False).encode('utf-8')

    resp = requests.post(url, data=json_data, headers={'Content-Type': 'application/json; charset=utf-8'}, timeout=30)
    data = resp.json()

    errcode = data.get("errcode", 0)
    if errcode != 0:
        raise Exception(f"获取用户数据失败: {data}")

    list_data = data.get("list", [])
    if not list_data:
        print(f"\n[用户增减数据] {begin_date} ~ {end_date}，无数据")
        return data

    # 表头
    print(f"\n[用户增减数据] {begin_date} ~ {end_date}")
    print(f"  {'日期':<12} {'新增用户':>8} {'取消关注':>8} {'净增关注':>8} {'累计关注':>10}")
    print(f"  {'-'*12} {'-'*8} {'-'*8} {'-'*8} {'-'*10}")

    total_new = 0
    total_cancel = 0
    for item in list_data:
        ref_date = item.get("ref_date", "")
        new_user = item.get("new_user", 0)
        cancel_user = item.get("cancel_user", 0)
        net = new_user - cancel_user
        cumulate = item.get("cumulate_user", "-")
        total_new += new_user
        total_cancel += cancel_user
        net_str = f"+{net}" if net >= 0 else str(net)
        cumulate_str = str(cumulate) if cumulate != "-" else "-"
        print(f"  {ref_date:<12} {new_user:>8} {cancel_user:>8} {net_str:>8} {cumulate_str:>10}")

    total_net = total_new - total_cancel
    total_net_str = f"+{total_net}" if total_net >= 0 else str(total_net)
    print(f"  {'='*12} {'='*8} {'='*8} {'='*8} {'='*10}")
    print(f"  {'合计':<12} {total_new:>8} {total_cancel:>8} {total_net_str:>8}")
    return data


def user_info(openid):
    """获取用户基本信息"""
    access_token = get_access_token()
    url = f"{API_USER_INFO}?access_token={access_token}&openid={openid}&lang=zh_CN"
    resp = requests.get(url, timeout=30)
    data = resp.json()

    errcode = data.get("errcode", 0)
    if errcode != 0:
        raise Exception(f"获取用户信息失败: {data}")

    subscribe = data.get("subscribe", 0)
    if not subscribe:
        print(f"\n[用户信息] openid: {openid}")
        print(f"  ⚠️ 该用户未关注公众号")
        return data

    nickname = data.get("nickname", "-")
    sex_map = {"0": "未知", "1": "男", "2": "女"}
    sex = sex_map.get(str(data.get("sex", "")), "-")
    province = data.get("province", "-")
    city = data.get("city", "-")
    country = data.get("country", "-")
    subscribe_time = data.get("subscribe_time", 0)
    subscribe_time_str = datetime.fromtimestamp(subscribe_time).strftime("%Y-%m-%d %H:%M") if subscribe_time else "-"
    groupid = data.get("groupid", "-")
    tagids = data.get("tagid_list", [])
    tagids_str = ", ".join(str(t) for t in tagids) if tagids else "无"
    remark = data.get("remark", "-") or "-"

    print(f"\n[用户基本信息]")
    print(f"  OpenID:   {openid}")
    print(f"  昵称:     {nickname}")
    print(f"  性别:     {sex}")
    print(f"  地区:     {country} {province} {city}")
    print(f"  关注时间: {subscribe_time_str}")
    print(f"  分组ID:   {groupid}")
    print(f"  标签:     {tagids_str}")
    print(f"  备注:     {remark}")
    return data


def user_list(next_openid=""):
    """获取用户列表"""
    access_token = get_access_token()
    url = f"{API_USER_LIST}?access_token={access_token}"
    if next_openid:
        url += f"&next_openid={next_openid}"

    resp = requests.get(url, timeout=30)
    data = resp.json()

    errcode = data.get("errcode", 0)
    if errcode != 0:
        raise Exception(f"获取用户列表失败: {data}")

    total = data.get("total", 0)
    count = data.get("count", 0)
    next_openid_out = data.get("next_openid", "")

    print(f"\n[用户列表]")
    print(f"  总关注人数: {total}")
    print(f"  本次返回:   {count}")
    print(f"  下一页ID:  {next_openid_out or '(无更多)'}")
    print()

    openids = data.get("data", {}).get("openid", [])
    for i, oid in enumerate(openids):
        print(f"  [{i+1}] {oid}")
    if not openids:
        print("  (空)")
    return data


def print_usage():
    print("""
微信公众号草稿推送工具 v2.0

用法:
  python wechat_push.py list                          查看草稿列表（含标题+更新时间）
  python wechat_push.py create <文件路径>              创建新草稿（支持 .html 和 .md）
  python wechat_push.py update <media_id> <文件路径> [--force-cover]   更新已有草稿（支持 .html 和 .md）
  python wechat_push.py delete <media_id>              删除草稿
  python wechat_push.py find <关键词>                  按标题关键词搜索草稿
  python wechat_push.py batch-del <id1> [id2] ...     批量删除草稿
  python wechat_push.py upload <图片文件>              上传永久素材（图片）
  python wechat_push.py materialcount                  获取永久素材总数统计
  python wechat_push.py materials [type] [count] [offset]  批量获取永久素材列表
                                                          type: image/video/voice/news，默认 image
  python wechat_push.py materialdel <media_id>         删除永久素材
  python wechat_push.py userstat [天数]               获取用户增减数据（默认近7天）
  python wechat_push.py userstat <begin> <end>         指定日期范围查询
  python wechat_push.py userinfo <openid>              获取用户基本信息
  python wechat_push.py userlist [next_openid]         获取用户列表
  python wechat_push.py cover <标题> [html文件]        生成封面图（指定HTML则保存到同目录）
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
                print("[ERROR] 请指定文件路径（支持 .html 和 .md）")
                print("用法: python wechat_push.py create <文件路径> [--force-cover]")
                print("       支持: .html (直接推送) 或 .md (自动转HTML后推送)")
                sys.exit(1)
            html_path = args[1]
            remaining = args[2:]
            force_cover = '--force-cover' in remaining
            draft_create(html_path, force_cover=force_cover)

        elif cmd == 'update':
            if len(args) < 2:
                print("[ERROR] 请指定 media_id 和文件路径")
                print("用法: python wechat_push.py update <media_id> <文件路径> [--force-cover]")
                print("       支持: .html (直接推送) 或 .md (自动转HTML后推送)")
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

        elif cmd == 'find':
            if len(args) < 2:
                print("[ERROR] 请指定搜索关键词")
                print("用法: python wechat_push.py find <关键词>")
                sys.exit(1)
            draft_find(args[1])

        elif cmd == 'batch-del':
            if len(args) < 2:
                print("[ERROR] 请指定要删除的 media_id（至少一个）")
                print("用法: python wechat_push.py batch-del <id1> [id2] [id3] ...")
                sys.exit(1)
            draft_batch_del(args[1:])

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

        elif cmd == 'userstat':
            # 解析日期范围，支持 userstat / userstat 7 / userstat 2026-03-01 2026-03-07
            from datetime import datetime, timedelta
            today = datetime.now()
            if len(args) >= 2 and '-' in args[1]:
                begin_date = args[1]
                end_date = args[2] if len(args) >= 3 else args[1]
            else:
                days = int(args[1]) if len(args) >= 2 else 7
                end_date = (today).strftime("%Y-%m-%d")
                begin_date = (today - timedelta(days=days - 1)).strftime("%Y-%m-%d")
            user_summary(begin_date, end_date)

        elif cmd == 'userinfo':
            if len(args) < 2:
                print("[ERROR] 请指定 OpenID")
                print("用法: python wechat_push.py userinfo <openid>")
                sys.exit(1)
            user_info(args[1])

        elif cmd == 'userlist':
            next_openid = args[1] if len(args) >= 2 else ""
            user_list(next_openid)

        elif cmd == 'cover':
            if len(args) < 2:
                print("[ERROR] 请指定标题")
                print("用法: python wechat_push.py cover <标题> [html文件]")
                sys.exit(1)
            title = args[1]
            if len(args) >= 3 and args[2].endswith('.html'):
                # 指定了HTML文件，封面保存到同目录
                html_path = args[2]
                title = args[1]
                dir_path = os.path.dirname(os.path.abspath(html_path))
                safe_name = re.sub(r'[^\w\u4e00-\u9fff]', '_', title)[:20]
                output_path = os.path.join(dir_path, f"cover_{safe_name}.png")
            else:
                # 未指定HTML，保存到 skill 目录下的 TMP/
                script_dir = os.path.dirname(os.path.abspath(__file__))
                tmp_dir = os.path.join(script_dir, "TMP")
                os.makedirs(tmp_dir, exist_ok=True)
                safe_name = re.sub(r'[^\w\u4e00-\u9fff]', '_', title)[:20]
                output_path = os.path.join(tmp_dir, f"cover_{safe_name}.png")

            result = generate_cover(title, output_path)
            if result:
                print(f"[OK] 封面图已生成: {output_path}")

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
