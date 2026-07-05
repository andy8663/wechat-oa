import pytest
from pathlib import Path


def test_cover_generator():
    from wechat_oa.features import generate_cover
    
    output_path = generate_cover("测试标题")
    assert Path(output_path).exists()
    Path(output_path).unlink()
