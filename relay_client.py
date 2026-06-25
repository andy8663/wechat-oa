#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wechat-oa relay client
中转模式客户端：将文章推送到公网服务器，由服务器转发至微信公众号平台
"""

import sys
import os
import json
import base64
import requests
from pathlib import Path

CONFIG_FILE = Path(__file__).parent / "config.json"


def load_config():
    """从 config.json 加载配置"""
    default_config = {
        "APP_ID": "",
        "APP_SECRET": "",
        "author": "Woody",
        "PUSH_MODE": "direct",
        "RELAY_SERVER": "http://120.79.2.44:8000",
        "WECHAT_OA_SERVER_KEY": "",
    }
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                user_config = json.load(f)
            return {**default_config, **user_config}
        except Exception:
            return default_config
    return default_config


def _post(url: str, payload: dict, api_key: str, timeout: int = 30) -> dict:
    """POST 请求封装，统一错误处理"""
    headers = {
        "Content-Type": "application/json",
        "X-API-Key": api_key,
    }
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=timeout)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.RequestException as e:
        return {"success": False, "error": f"网络请求失败: {e}"}
    except json.JSONDecodeError:
        return {"success": False, "error": f"服务器返回非 JSON 数据: {resp.text[:200]}"}


# ─────────────────────────────────────────────────────────────────────────────
# 一站式推送：标题 + 正文 + 封面图 → 服务器 → 公众号草稿箱
# ─────────────────────────────────────────────────────────────────────────────

def push_article(title: str, content: str, author: str = "", digest: str = "",
                 thumb_path: str = None, api_key: str = "", relay_server: str = "") -> dict:
    """
    一站式推送文章到公众号草稿箱（通过中转服务器）
    
    Args:
        title: 文章标题
        content: 文章正文 HTML
        author: 作者
        digest: 摘要
        thumb_path: 封面图本地路径（可选，会 base64 编码后发送）
        api_key: WECHAT_OA_SERVER_KEY
        relay_server: 中转服务器地址（如 http://120.79.2.44:8000）
    
    Returns:
        dict: {"success": True/False, "media_id": "...", "message": "..."}
    """
    cfg = load_config()
    if not api_key:
        api_key = cfg.get("WECHAT_OA_SERVER_KEY", "")
    if not relay_server:
        relay_server = cfg.get("RELAY_SERVER", "http://120.79.2.44:8000")

    if not api_key:
        return {"success": False, "error": "未配置 WECHAT_OA_SERVER_KEY"}

    # 构建请求体
    payload = {
        "appid": cfg.get("APP_ID", ""),
        "appsecret": cfg.get("APP_SECRET", ""),
        "title": title,
        "content": content,
        "author": author or cfg.get("author", "Woody"),
        "digest": digest,
    }

    # 封面图：读取并 base64 编码
    if thumb_path and os.path.exists(thumb_path):
        try:
            with open(thumb_path, 'rb') as f:
                img_data = f.read()
            payload["thumb_image"] = base64.b64encode(img_data).decode('utf-8')
            payload["thumb_filename"] = os.path.basename(thumb_path)
        except Exception as e:
            print(f"[WARN] 封面图读取失败，将继续无封面推送: {e}")

    url = f"{relay_server.rstrip('/')}/api/push/article"
    result = _post(url, payload, api_key)

    if result.get("success"):
        return {"success": True, "media_id": result.get("media_id", ""), "message": result.get("message", "")}
    else:
        return {"success": False, "error": result.get("error", "未知错误")}


# ─────────────────────────────────────────────────────────────────────────────
# 列出草稿箱（通过中转服务器）
# ─────────────────────────────────────────────────────────────────────────────

def list_drafts(count: int = 10, offset: int = 0, api_key: str = "", relay_server: str = "") -> dict:
    """
    列出公众号草稿箱（通过中转服务器）
    
    Args:
        count: 拉取数量（最大 20）
        offset: 偏移量
        api_key: WECHAT_OA_SERVER_KEY
        relay_server: 中转服务器地址
    
    Returns:
        dict: {"success": True/False, "drafts": [...], "total": N}
    """
    cfg = load_config()
    if not api_key:
        api_key = cfg.get("WECHAT_OA_SERVER_KEY", "")
    if not relay_server:
        relay_server = cfg.get("RELAY_SERVER", "http://120.79.2.44:8000")

    if not api_key:
        return {"success": False, "error": "未配置 WECHAT_OA_SERVER_KEY"}

    params = {
        "appid": cfg.get("APP_ID", ""),
        "appsecret": cfg.get("APP_SECRET", ""),
        "count": min(count, 20),
        "offset": offset,
    }

    url = f"{relay_server.rstrip('/')}/api/push/drafts"
    headers = {"X-API-Key": api_key}
    try:
        resp = requests.get(url, headers=headers, params=params, timeout=15)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.RequestException as e:
        return {"success": False, "error": f"网络请求失败: {e}"}
    except json.JSONDecodeError:
        return {"success": False, "error": f"服务器返回非 JSON 数据"}


# ─────────────────────────────────────────────────────────────────────────────
# CLI 入口（供命令行直接调用测试）
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="wechat-oa relay client")
    subparsers = parser.add_subparsers(dest="command")

    # push 子命令
    push_parser = subparsers.add_parser("push", help="推送文章到公众号草稿箱")
    push_parser.add_argument("html_path", help="HTML 文件路径")
    push_parser.add_argument("--title", default="", help="文章标题")
    push_parser.add_argument("--author", default="", help="作者")
    push_parser.add_argument("--digest", default="", help="摘要")
    push_parser.add_argument("--thumb", default="", help="封面图路径")

    # list 子命令
    list_parser = subparsers.add_parser("list", help="列出草稿箱")
    list_parser.add_argument("--count", type=int, default=10, help="拉取数量")
    list_parser.add_argument("--offset", type=int, default=0, help="偏移量")

    args = parser.parse_args()

    if args.command == "push":
        if not os.path.exists(args.html_path):
            print(f"[ERROR] 文件不存在: {args.html_path}")
            sys.exit(1)
        with open(args.html_path, 'r', encoding='utf-8') as f:
            content = f.read()
        title = args.title or os.path.basename(args.html_path).replace('.html', '')
        result = push_article(title, content, args.author, args.digest, args.thumb or None)
        print(json.dumps(result, ensure_ascii=False, indent=2))

    elif args.command == "list":
        result = list_drafts(args.count, args.offset)
        print(json.dumps(result, ensure_ascii=False, indent=2))

    else:
        parser.print_help()
