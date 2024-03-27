import lark_oapi as lark
import json
import pandas as pd
from os import environ
from datetime import datetime
from dotenv import load_dotenv
from lark_oapi.api.task.v2 import *

load_dotenv()
user_Access_Token = environ.get("USER_ACCESS_TOKEN")
task_Guid = environ.get("TASK_GUID")

def main():
    client = lark.Client.builder() \
        .enable_set_token(True) \
        .log_level(lark.LogLevel.DEBUG) \
        .build()

    request: GetTaskRequest = GetTaskRequest.builder() \
        .task_guid(task_Guid) \
        .user_id_type("open_id") \
        .build()

    option = lark.RequestOption.builder().user_access_token(user_Access_Token).build()
    response: GetTaskResponse = client.task.v2.task.get(request, option)

    if not response.success():
        lark.logger.error(
            f"client.task.v2.task.get failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
        return

    # 处理业务结果
    #lark.logger.info(lark.JSON.marshal(response.data, indent=4))

    json_str = lark.JSON.marshal(response.data, indent=4)
    data = json.loads(json_str)
    custom_fields = data['task']['custom_fields']
    
    df = pd.DataFrame(custom_fields)
    df = df.assign(序号=df['guid'])
    
    df.to_excel('output.xlsx', index=False)

    print("ok")

def TimeChange(unixTime):
    timestamp = unixTime / 1000
    return datetime.fromtimestamp(timestamp)

if __name__ == "__main__":
    main()