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

    # header_code 删除表中所有数据 获取连接池
    with closing(get_connection()) as conn:
        with conn.cursor() as cursor:
            cursor.execute('DELETE FROM bug_data')
            conn.commit()

    header_code = header_code_return()[0]
    len_header_code = len(header_code)
    # 测试
    for header_item in range(len_header_code):
        print('header_item', header_item)
        next_token1 = ''
        header_code_result = header_code[header_item]

        first_request = ali_bug_query.main(next_token1, header_code_result)
        total_count = first_request.body.total_count
        count = int(total_count)

        project = ''
        next_token = ''
        for i in range(count):
            request_data = ali_bug_query.main(next_token, header_code[header_item])
            next_token = request_data.body.next_token
            max_results = len(request_data.body.workitems)
            for y in range(max_results):
                print('y----', y)
                print('max_results---', max_results)
                print('request_data-----', request_data)
                print('next_token-----', next_token)
                print('first_request-----', first_request)
                print('count-----', count)
                print('i-----', i)

                batch_data = []
                item = request_data.body.workitems[y]
                detail_link = f'https://devops.aliyun.com/projex/project/{item.space_identifier}/bug/{item.identifier}'
                row = ((
                    item.assigned_to, item.category_identifier, item.creator,
                    item.document, item.gmt_create, item.gmt_modified,
                    item.identifier, item.logical_status, item.modifier,
                    item.parent_identifier, item.serial_number, item.space_identifier,
                    item.space_name, item.space_type, item.sprint_identifier,
                    item.status, item.status_identifier, item.status_stage_identifier,
                    item.subject, item.workitem_type_identifier, detail_link
                ))

                batch_data.append(row)

                with closing(get_connection()) as conn:
                    with conn.cursor() as cursor:
                        cursor.executemany('''
                            INSERT INTO bug_data (
                                assignedTo, categoryIdentifier, creator, document,
                                gmtCreate, gmtModified, identifier, logicalStatus,
                                modifier, parentIdentifier, serialNumber, spaceIdentifier,
                                spaceName, spaceType, sprintIdentifier, status,
                                statusIdentifier, statusStageIdentifier, subject,
                                workitemTypeIdentifier, detail_link
                            ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                        ''', batch_data)
                        conn.commit()
                        print('query_bug_data '+ str(header_item)+ ' '+str(project) +' ' + str(next_token)+ ' '+str(i*200+y))




