import os
import json
from dotenv import load_dotenv

# 加载 .env 文件中的环境变量
load_dotenv()

def interpolate_env_vars(config):
    """递归地替换配置中的环境变量占位符"""
    if isinstance(config, dict):
        return {key: interpolate_env_vars(value) for key, value in config.items()}
    elif isinstance(config, list):
        return [interpolate_env_vars(item) for item in config]
    elif isinstance(config, str):
        # 替换 ${VAR_NAME} 格式的环境变量
        import re
        pattern = r'\$\{([^}]+)\}'
        def replace_var(match):
            var_name = match.group(1)
            # 优先使用配置中的值，然后是环境变量，最后是 .env 文件中的值
            return os.environ.get(var_name, '')
        return re.sub(pattern, replace_var, config)
    else:
        return config

def load_and_interpolate_config(config_file_path):
    """加载并插值配置文件"""
    with open(config_file_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    # 插值环境变量
    interpolated_config = interpolate_env_vars(config)
    
    return interpolated_config