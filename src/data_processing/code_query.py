import pymysql
from config.database import read_db_config

def query_code_date(year, month):
    db_config = read_db_config()
    conn = pymysql.connect(**db_config)
    cursor = conn.cursor()
    # 执行查询
    cursor.execute('SELECT * FROM code_data WHERE year=%s AND month=%s;', (year, month))
    results = cursor.fetchall()
    # 处理结果
    cursor.close()
    conn.close()
    return results

def query_code_bug_date(year, month):
    pass