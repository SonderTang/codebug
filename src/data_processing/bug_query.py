import pymysql
from config.database import read_db_config
from config.database import get_connection
from contextlib import closing

def header_code_return():
    spaceName_list = []
    space_list = []
    with closing(get_connection()) as conn: # 自动归还连接
        with conn.cursor() as cursor:
            cursor.execute('SELECT * FROM spacename;')

            while True:
                res = cursor.fetchall()
                if not res:
                    break

                spaceName_list.append(res[1])
                space_list.append(res[2])

    return spaceName_list, space_list

def query_bug_data():
    # db_config = read_db_config()
    # conn = pymysql.connect(**db_config)
    # cursor = conn.cursor()
    # cursor.execute('delete from bug_data')
    # conn.commit()

    # header_code
    with closing(get_connection()) as conn:
        with conn.cursor() as cursor:
            cursor.execute('DELETE FROM bug_data')
            conn.commit()

    header_code = header_code_return()[0]


