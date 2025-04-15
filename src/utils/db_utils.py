import pymysql
from config.database import read_db_config

def get_db_connection():
    db_config = read_db_config()
    return pymysql.connect(**db_config)