import json
import os
from pathlib import Path
from typing import Dict, Any, Optional

DEFAULT_CONFIG = {
    "APP_ID": "",
    "APP_SECRET": "",
    "author": "",
    "PUSH_MODE": "direct",
    "WECHAT_OA_SERVER": "http://120.79.2.44",
    "WECHAT_OA_SERVER_KEY": "",
    "ENV": "prod"
}

_config_cache: Optional[Dict[str, Any]] = None


def load_config(config_path: str = "") -> Dict[str, Any]:
    global _config_cache
    
    if _config_cache is not None:
        return _config_cache
    
    if config_path:
        config_file = Path(config_path)
    else:
        config_file = Path(__file__).parent.parent.parent / "config.json"
    
    config = DEFAULT_CONFIG.copy()
    
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                user_config = json.load(f)
            config.update(user_config)
        except Exception as e:
            raise ConfigError(f"加载配置文件失败: {e}")
    
    env_config = {k: v for k, v in os.environ.items() if k.startswith("WECHAT_OA_")}
    for key, value in env_config.items():
        config_key = key.replace("WECHAT_OA_", "")
        if config_key in config:
            config[config_key] = value
    
    _config_cache = config
    return config


def validate_config(config: Dict[str, Any]) -> None:
    if not config.get("APP_ID"):
        raise ConfigError("APP_ID 不能为空")
    if not config.get("APP_SECRET"):
        raise ConfigError("APP_SECRET 不能为空")
    if config.get("PUSH_MODE") not in ("direct", "relay", "hybrid"):
        raise ConfigError(f"无效的 PUSH_MODE: {config.get('PUSH_MODE')}")


class ConfigError(Exception):
    pass
