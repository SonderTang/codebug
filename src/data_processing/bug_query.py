import pymysql
from config.database import read_db_config
from config.database import get_connection
from contextlib import closing
from src.ali.get_bug_list import ali_bug_query


def header_code_return():
    spaceName_list = []
    space_list = []
    with closing(get_connection()) as conn: # 自动归还连接
        with conn.cursor() as cursor:
            cursor.execute('SELECT * FROM spacename;')
            for result in cursor:
                print(result)
                space_list.append(result.get('space'))
                spaceName_list.append(result.get('spaceName'))
    print('spaceName_list', spaceName_list)
    print('space_list', space_list)
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

    # header_code_return(['2f47d6d1e8613e642d7abe6d99', '403c693bcbbb4e0af420b62f57', '75f7ce1c6cf94901b4f322ad22',
    #                     '4eda1e506a21f3445c17ec0a9c', '76be871b30aac2b237a3657a83', 'd025ff035475f72da623dbaacc',
    #                     'dbf8ae0e00a6d48a519088c486', '49495e17b3882ca0217fa920f7', '1149ac502df9de5291b17a4240',
    #                     'ba3ba53c629658eb964a3e6c30', '9f78454755955a027b9593ca51', 'c71cd7bba19a51b7898e8c1b15',
    #                     '59fb1b22e3a43f665ba993db6f', 'c0d2e0040c43d535656f5cb649'],
    #                    ['ERP', '独立站项目组', '数据中台', '权限认证中心', 'OA', 'iBay', 'MRP', '扬腾仓储', '前端技术项目', 'TMS', 'OMS',
    #                     '业财融合项目', '采购SRM项目', 'BPM'])

    header_code = header_code_return()[0]
    len_header_code = len(header_code)
    # 测试
    for header_item in range(len_header_code):
        print('header_item', header_item)
        next_token = ''
        header_code_result = header_code[header_item]

        first_request = ali_bug_query.main(next_token, header_code_result)
        total_count = first_request.body.total_count
        cout = int(total_count)
        print('first_request', first_request)
        print('total_count', total_count)
        print('cout', cout)

