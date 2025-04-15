import os

# 获取数据库配置
def read_api_config() -> dict:
    return {
        "id": os.getenv("API_KEY_ID"),
        "secret": os.getenv("API_KEY_SECRET"),
        "origin_id": os.getenv("API_ORIGIN_ID")
    }