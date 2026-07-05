import pytest
from wechat_oa.convert import get_converter, MdToWechatConverter, HtmlToWechatConverter
from wechat_oa.core.utils import count_wechat_units, truncate_digest, validate_digest, extract_digest


def test_count_wechat_units():
    assert count_wechat_units("hello") == 2.5
    assert count_wechat_units("你好") == 2.0
    assert count_wechat_units("你好hello") == 4.5


def test_truncate_digest():
    long_text = "这是一个非常非常非常非常非常非常非常非常非常非常非常非常非常非常长的摘要"
    truncated = truncate_digest(long_text, 10)
    assert count_wechat_units(truncated) <= 10


def test_validate_digest():
    digest, truncated = validate_digest("短摘要")
    assert not truncated
    assert digest == "短摘要"
    
    long_digest = "这是一个非常长的摘要" * 20
    digest, truncated = validate_digest(long_digest)
    assert truncated
    assert count_wechat_units(digest) <= 120


def test_extract_digest():
    html = "<p>这是第一段有意义的内容，包含了文章的核心观点。</p><p>这是第二段内容。</p>"
    digest = extract_digest(html)
    assert len(digest) > 0
    assert count_wechat_units(digest) <= 120


def test_get_converter():
    md_conv = get_converter("test.md")
    assert isinstance(md_conv, MdToWechatConverter)
    
    html_conv = get_converter("test.html")
    assert isinstance(html_conv, HtmlToWechatConverter)
    
    with pytest.raises(ValueError):
        get_converter("test.txt")
