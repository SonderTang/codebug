import os
import pymysql
from dbutils.pooled_db import PooledDB
from pymysql.cursors import DictCursor

# 连接池全局单例初始化
db_pool = PooledDB(
    creator=pymysql,
    mincached=2,          # 初始空闲连接数
    maxcached=5,          # 最大空闲连接数（0=无限制）
    maxconnections=10,    # 总连接数上限
    blocking=True,        # 连接耗尽时阻塞等待而非报错
    host=os.getenv("DB_HOST"),
    port=int(os.getenv("DB_PORT", "3306")),
    user=os.getenv("DB_USER"),
    password=os.getenv("DB_PASSWORD"),
    database=os.getenv("DB_NAME"),
    charset='utf8',
    cursorclass=DictCursor,
    ping=1                # 每次取连接时检查活性
)

def get_connection():
    return db_pool.connection()

# 获取数据库配置
def read_db_config() -> dict:
    return {
        "host": os.getenv("DB_HOST"),
        "port": int(os.getenv("DB_PORT", "3306")),
        "user": os.getenv("DB_USER"),
        "password": os.getenv("DB_PASSWORD"),
        "database": os.getenv("DB_NAME"),
        "charset": os.getenv("DB_CHARSET", "utf8")
    }




