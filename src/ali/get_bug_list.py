from typing import List
from alibabacloud_devops20210625.client import Client as devops20210625Client
from alibabacloud_tea_openapi import models as open_api_models
from alibabacloud_devops20210625 import models as devops_20210625_models
from alibabacloud_tea_util import models as util_models
from alibabacloud_tea_util.client import Client as UtilClient
from config.database import read_db_config

from config.api_keys import read_api_config

class ali_bug_query:
    def __init__(self):
        pass

    @staticmethod
    def create_client() -> devops20210625Client:
        api_config = read_api_config()
        config = open_api_models.Config(
            access_key_id=api_config.get('id'),
            access_key_secret=api_config.get('secret')
        )
        config.endpoint = f'devops.cn-hangzhou.aliyuncs.com'
        return devops20210625Client(config)

    @staticmethod
    def main(next_token, space_identifier):
        api_config = read_api_config()
        client = ali_bug_query.create_client()
        list_workitems_request = devops_20210625_models.ListWorkitemsRequest(
            space_type='Project',
            space_identifier=space_identifier,
            category='Bug',
            next_token=next_token,
            max_results = '200',
            conditions='{"conditionGroups": [[{"fieldIdentifier": "gmtCreate","operator": "MORE_THAN_AND_EQUAL","value": ["2025-03-01 00:00:00"]}]]}',
        )
        runtime = util_models.RuntimeOptions()
        headers = {}
        try:
            # 复制代码运行请自行打印 API 的返回值
            ListWorkitemsResponse = client.list_workitems_with_options(api_config.get('origin_id'), list_workitems_request, headers, runtime)
            # a =  json.dumps(ListWorkitemsResponse)
            return ListWorkitemsResponse
        except Exception as error:
            # 如有需要，请打印 error
            UtilClient.assert_as_string(error)
            print(UtilClient.assert_as_string(error))