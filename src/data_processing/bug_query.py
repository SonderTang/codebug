import time

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

    # 循环所有项目
    for header_item in range(len_header_code):
        print('header_item', header_item)
        next_token1 = ''
        header_code_result = header_code[header_item]

        request_data = ali_bug_query.main(next_token1, header_code_result)
        total_count = request_data.body.total_count
        count = int(total_count)

        project = ''
        next_token = ''

        # 在单个项目内
        # request_data = ali_bug_query.main(next_token, header_code[header_item])
        next_token = request_data.body.next_token
        max_results = len(request_data.body.workitems)
        for y in range(max_results):
            batch_data = []
            item = request_data.body.workitems[y]
            detail_link = f'https://devops.aliyun.com/projex/project/{item.space_identifier}/bug/{item.identifier}'

            with closing(get_connection()) as conn:
                with conn.cursor() as cursor:
                    cursor.execute('select id from bug_data order by id desc limit 1;')
                    ids = cursor.fetchone()
                    if ids is None:
                        id = 1
                    else:
                        id = ids['id'] + 1
                    row = ((
                        id,
                        item.assigned_to, item.category_identifier, item.creator,
                        item.document, item.gmt_create, item.gmt_modified,
                        item.identifier, item.logical_status, item.modifier,
                        item.parent_identifier, item.serial_number, item.space_identifier,
                        item.space_name, item.space_type, item.sprint_identifier,
                        item.status, item.status_identifier, item.status_stage_identifier,
                        item.subject, item.workitem_type_identifier, detail_link
                    ))

                    batch_data.append(row)
                    cursor.executemany('''
                                    INSERT INTO bug_data (
                                        id,
                                        assignedTo, categoryIdentifier, creator, document,
                                        gmtCreate, gmtModified, identifier, logicalStatus,
                                        modifier, parentIdentifier, serialNumber, spaceIdentifier,
                                        spaceName, spaceType, sprintIdentifier, status,
                                        statusIdentifier, statusStageIdentifier, subject,
                                        workitemTypeIdentifier, detail_link
                                    ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                                ''', batch_data)
                    conn.commit()
                    print('header_code_result', header_code_result)

    # 更新bug_data中的一些字段
    with closing(get_connection()) as conn:
        with conn.cursor() as cursor:
            cursor.execute('select count(1) AS total from bug_data') # bug总数量
            bug_total = cursor.fetchone()['total']
            print('total', bug_total)
            for m in range(bug_total):
               n = m + 1
               cursor.execute('select creator from bug_data where id = %s'%n)
               creator_id = cursor.fetchone()['creator']
               cursor.execute('select realName from user_data where identifier = %s;'%creator_id)
               realName_result = cursor.fetchone()
               if realName_result is not None:
                   realName = realName_result['realName']
                   cursor.execute('update bug_data set creator = %s where id = %s', (realName, n))
                   conn.commit()

               cursor.execute('select gmtCreate,gmtModified from bug_data where id = %s'%n)
               date_result = cursor.fetchone()
               gmtCreate_1 = int(date_result['gmtCreate']) / 1000
               gmtModified_1 = int(date_result['gmtModified']) / 1000
               gmtCreate_2 = time.localtime(gmtCreate_1)
               gmtModified_2 = time.localtime(gmtModified_1)
               gmtCreate = time.strftime("%Y-%m-%d %H:%M:%S", gmtCreate_2)
               gmtModified = time.strftime("%Y-%m-%d %H:%M:%S", gmtModified_2)
               cursor.execute('update bug_data set gmtCreate= %s,gmtModified=%s where id = %s;',
                            (gmtCreate_1, gmtModified_1, n))
               conn.commit()
               cursor.execute('select modifier from bug_data where id = %s' % n)
               creator_id = cursor.fetchone()['modifier']
               if creator_id is not None:
                   cursor.execute('select realName from user_data where identifier = %s;'%creator_id)
                   realName_result = cursor.fetchone()
                   if realName_result is not None:
                       realName = realName_result['realName']
                       cursor.execute('update bug_data set modifier = %s where id = %s;', (realName, n))
                       conn.commit()

            # 更新bug_data的assignedTo字段值，负责人
            cursor.execute('select * from bug_data;')
            bug_result = cursor.fetchall()
            len_bug_result = len(bug_result)
            for o in range(len_bug_result):
                assignedTo = bug_result[o]['assignedTo']
                id_bug = bug_result[o]['id']
                cursor.execute('select realName from user_data where identifier = %s;' % assignedTo)
                realName = cursor.fetchone()
                if realName is None:
                    realName_result = '查无此人'
                else:
                    realName_result = realName['realName']
                cursor.execute('update bug_data set assignedTo = %s where id =%s;', (realName_result, id_bug))
                conn.commit()



