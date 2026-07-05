import pytest
from wechat_oa.core.config import load_config, DEFAULT_CONFIG


def test_load_config_default():
    config = load_config()
    for key in DEFAULT_CONFIG:
        assert key in config


def test_config_validation():
    from wechat_oa.core.config import validate_config, ConfigError
    
    valid_config = {
        "APP_ID": "test_id",
        "APP_SECRET": "test_secret",
        "PUSH_MODE": "direct"
    }
    validate_config(valid_config)
    
    with pytest.raises(ConfigError):
        validate_config({"APP_ID": "", "APP_SECRET": "test"})
    
    with pytest.raises(ConfigError):
        validate_config({"APP_ID": "test", "APP_SECRET": ""})
    
    with pytest.raises(ConfigError):
        validate_config({"APP_ID": "test", "APP_SECRET": "test", "PUSH_MODE": "invalid"})
