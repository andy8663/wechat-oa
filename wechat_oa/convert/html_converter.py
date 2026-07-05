import re
from pathlib import Path
from typing import Dict


class HtmlToWechatConverter:
    ALLOWED_TAGS = [
        'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
        'strong', 'b', 'em', 'i', 'u', 's', 'del',
        'a', 'br', 'hr', 'blockquote',
        'ul', 'ol', 'li', 'table', 'tbody', 'tr', 'td',
        'img', 'div', 'span', 'section'
    ]
    
    ALLOWED_ATTRIBUTES = {
        'a': ['href', 'target'],
        'img': ['src', 'alt', 'width', 'height'],
        '*': ['style']
    }
    
    def convert(self, html_path: str, user_digest: str = "") -> Dict:
        from wechat_oa.core.utils import extract_digest
        
        with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
            html = f.read()
        
        try:
            from premailer import Premailer
            p = Premailer(
                html,
                remove_classes=False,
                strip_important=False,
                include_star_selectors=False,
                disable_link_rewrites=True
            )
            html = p.transform()
        except ImportError:
            pass
        
        try:
            import bleach
            html = bleach.clean(
                html,
                tags=self.ALLOWED_TAGS,
                attributes=self.ALLOWED_ATTRIBUTES,
                strip=True
            )
        except ImportError:
            pass
        
        title = self._extract_title(html)
        body = self._extract_body(html)
        
        return {
            "title": title,
            "body": body,
            "digest": user_digest if user_digest else extract_digest(body),
            "author": ""
        }
    
    def _extract_title(self, html: str) -> str:
        m = re.search(r'<title>([^<]+)</title>', html, re.IGNORECASE)
        if m:
            return m.group(1).strip()[:64]
        
        m = re.search(r'<h1[^>]*>([^<]+)</h1>', html, re.IGNORECASE)
        if m:
            return m.group(1).strip()[:64]
        
        return "无标题"
    
    def _extract_body(self, html: str) -> str:
        m = re.search(r'<body[^>]*>(.*)</body>', html, re.DOTALL | re.IGNORECASE)
        if m:
            return m.group(1).strip()
        
        return html.strip()
