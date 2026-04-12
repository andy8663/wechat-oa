#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import re
sys.stdout.reconfigure(encoding='utf-8')

def parse_html_article(html_path):
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

    # 微信公众号标题限制：最多64个字符
    if len(title) > 64:
        title = title[:64]

    # 提取正文
    body_match = re.search(r'<body[^>]*>(.*?)</body>', content, re.IGNORECASE | re.DOTALL)
    if body_match:
        body = body_match.group(1)
    else:
        html_match = re.search(r'<html[^>]*>(.*?)</html>', content, re.IGNORECASE | re.DOTALL)
        if html_match:
            body = html_match.group(1)
        else:
            body = content
        body = re.sub(r'<head[^>]*>.*?</head>', '', body, flags=re.IGNORECASE | re.DOTALL)
        body = re.sub(r'<title[^>]*>.*?</title>', '', body, flags=re.IGNORECASE | re.DOTALL)

    body = re.sub(r'<script[^>]*>.*?</script>', '', body, flags=re.IGNORECASE | re.DOTALL)
    body = re.sub(r'</li>\s*\n\s*<li', '</li><li', body, flags=re.IGNORECASE)
    body = re.sub(r'</li>\s*\n\s*</ul>', '</li></ul>', body, flags=re.IGNORECASE)
    body = re.sub(r'</li>\s*\n\s*</ol>', '</li></ol>', body, flags=re.IGNORECASE)
    body = re.sub(r'<ul[^>]*>\s*\n\s*<li', lambda m: m.group(0).replace('\n', '').replace('  ', ''), body, flags=re.IGNORECASE)
    body = re.sub(r'<ol[^>]*>\s*\n\s*<li', lambda m: m.group(0).replace('\n', '').replace('  ', ''), body, flags=re.IGNORECASE)
    body = re.sub(r'[ \t]+', ' ', body)
    body = re.sub(r'\n{2,}', '\n', body)
    body = re.sub(r' *\n *', '\n', body)
    body = re.sub(r'>\n+<', '><', body)

    return title.strip(), body.strip()

html_path = 'C:/Users/andy8/.qclaw/workspace/skills/wechat-mp-push/article-multi-agent-v12.html'
title, body = parse_html_article(html_path)

print('=== Parsed Title ===')
print(f'Length: {len(title)}')
print(f'Title: [{repr(title)}]')
print(f'is empty: {not title}')
print(f'is whitespace: {not title.strip()}')
print(f'first 100 chars of body: {body[:100]}')
