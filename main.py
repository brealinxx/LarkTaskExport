import lark_oapi as lark
import json
import pandas as pd
import sys
import os
import time
import warnings
import webbrowser
from os import environ
from datetime import datetime
from dotenv import load_dotenv
from lark_oapi.api.task.v2 import *
from lark_oapi.api.contact.v3 import *
from lark_oapi.api.authen.v1 import *
from lark_oapi.api.auth.v3 import *
from urllib.parse import urlparse, parse_qs
from contextlib import contextmanager

# todo 1.auto useraccesstoken(complete) 2.multithreading 2.hidden subtask

warnings.filterwarnings("ignore", category=FutureWarning)

if getattr(sys, "frozen", False):
    executable_dir = os.path.dirname(sys.executable)
    env_path = os.path.join(executable_dir, ".env")
else:
    script_dir = os.path.dirname(sys.argv[0])
    env_path = os.path.join(script_dir, ".env")
load_dotenv(env_path)
app_Id = environ.get("APP_ID")
app_Secret = environ.get("APP_SECRET")
tasklist_Guid = environ.get("TASKLIST_GUID")


@contextmanager
def get_redirect_url():
    redirect_url = (
        f"https://open.feishu.cn/open-apis/authen/v1/index?app_id={app_Id}&redirect_uri=https://open.feishu.cn/app/{app_Id}/safe&state="
        ""
    )
    webbrowser.open(redirect_url)

    redirected_url = input(
        "请复制登陆后的浏览器地址栏中的 URL 并粘贴在这里，然后按 Enter 键继续："
    )
    yield redirected_url


def GetCodeFromRedirectURL():
    with get_redirect_url() as redirected_url:
        parsed_url = urlparse(redirected_url)
        try:
            code = parse_qs(parsed_url.query).get("code")[0]
            return code
        except:
            raise ValueError("请检查登陆情况和 URL 是否复制正确并重新启动程序")


def initToken():
    if getattr(sys, "frozen", False):
        log_level = lark.LogLevel.INFO
    else:
        log_level = lark.LogLevel.DEBUG

    return (
        lark.Client.builder()
        .app_id(app_Id)
        .app_secret(app_Secret)
        .log_level(log_level)
        .build()
    )


code = GetCodeFromRedirectURL()


def GetUserAccessTokenRequest():
    client = initToken()
    request: CreateOidcAccessTokenRequest = (
        CreateOidcAccessTokenRequest.builder()
        .request_body(
            CreateOidcAccessTokenRequestBody.builder()
            .grant_type("authorization_code")
            .code(code)
            .build()
        )
        .build()
    )

    if request:
        response: CreateOidcAccessTokenResponse = (
            client.authen.v1.oidc_access_token.create(request)
        )
        if not response.success():
            lark.logger.error(
                f"client.authen.v1.oidc_access_token.create failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}"
            )
            return

        return response
    else:
        return


try:
    user_Access_Token = GetUserAccessTokenRequest().data.access_token
except:
    print("请检查 User_Access_Token")


def process_task_data(task_data, is_subtask=False):
    """
    从任务数据中提取所需信息并转换格式
    """
    if is_subtask:
        task_item = "  " + (task_data.summary or pd.NA)  # 在任务项内容前面添加两个空格
    else:
        task_item = task_data.summary or pd.NA

    processed_data = {
        "标题": GetSectionNameBySectionGuid(GetSectionGuid(task_data.tasklists))
        or pd.NA,
        "任务项": task_item,
        "创建人": GetNameByUserID(task_data.creator.id) or pd.NA,
        "任务创建时间": TimeChange(task_data.created_at) or pd.NA,
        "负责人": GetMemberNameByLoop(task_data.members) or pd.NA,
        "开始时间": TimeChange(task_data.start.timestamp) if task_data.start else pd.NA,
        "完成时间": TimeChange(task_data.completed_at) or pd.NA,
        "计划工时": GetCustomPlannedWorkingHoursFields(task_data.custom_fields)
        or pd.NA,
        "开发工时": GetCustomDevelopmentHoursFields(task_data.custom_fields) or pd.NA,
        "IsSubtask": is_subtask,
    }

    return processed_data


class LoadingTimer:
    def __init__(self):
        self.start_time = None
        self.elapsed_time = 0

    def __enter__(self):
        self.start_time = time.time()
        print("正在加载中...\n")
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        print("\n加载完成!")

    def update_timer(self):
        if self.start_time:
            self.elapsed_time = int(time.time() - self.start_time)
            sys.stdout.write(f"\r已加载 {self.elapsed_time} 秒")
            sys.stdout.flush()


def main():
    with LoadingTimer() as timer:
        tasks_response = GetTasksOfListRequest()
        tasks_data = tasks_response.data

        all_tasks_data = pd.DataFrame()

        for task in tasks_data.items:
            single_task_response = GetSingleTasksRequest(task.guid)
            subtasks_response = GetSubTaskOfListRequest(task.guid)
            try:
                single_task_data = single_task_response.data
                subtasks_data = subtasks_response.data
            except:
                print("请重置 USER_ACCESS_TOKEN 后重试")

            # 处理主任务数据
            processed_data = process_task_data(single_task_data.task, is_subtask=False)
            task_df = pd.DataFrame([processed_data])
            task_df = task_df.dropna(how="all", axis=1)

            all_tasks_data = pd.concat([all_tasks_data, task_df], ignore_index=True)

            # 处理子任务数据
            for subtask in subtasks_data.items:
                subtask_response = GetSingleTasksRequest(subtask.guid)
                subtask_data = subtask_response.data

                processed_subtask_data = process_task_data(
                    subtask_data.task, is_subtask=True
                )
                subtask_df = pd.DataFrame([processed_subtask_data])
                subtask_df = subtask_df.dropna(how="all", axis=1)
                all_tasks_data = pd.concat(
                    [all_tasks_data, subtask_df], ignore_index=True
                )

            timer.update_timer()

        # all_tasks_data = all_tasks_data.sort_values(by='任务创建时间', ascending=True)

        if getattr(sys, "frozen", False):
            file_path = executable_dir
        else:
            file_path = os.getcwd()
        excel_file_path = os.path.join(file_path, "output.xlsx")

        WriteToExcel(all_tasks_data, excel_file_path)


def WriteToExcel(df, excel_file_path):
    writer = pd.ExcelWriter(excel_file_path, engine="xlsxwriter")
    df.to_excel(writer, sheet_name="Sheet1", index=False)

    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]

    subtask_format = workbook.add_format({"font_size": 9})
    # worksheet.set_column('J:J', None, subtask_format, {'hidden': True})

    for idx, row in df.iterrows():
        if row["IsSubtask"]:
            worksheet.set_row(idx + 1, None, subtask_format)

    writer.close()


def init():
    if getattr(sys, "frozen", False):
        log_level = lark.LogLevel.INFO
    else:
        log_level = lark.LogLevel.DEBUG

    return lark.Client.builder().enable_set_token(True).log_level(log_level).build()


def GetSingleTasksRequest(task_Guid):
    client = init()
    request: GetTaskRequest = (
        GetTaskRequest.builder().task_guid(task_Guid).user_id_type("open_id").build()
    )

    option = lark.RequestOption.builder().user_access_token(user_Access_Token).build()
    response: GetTaskResponse = client.task.v2.task.get(request, option)
    if not response.success():
        lark.logger.error(
            f"client.task.v2.task.get failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}"
        )
        return

    return response


def GetUserNameRequest(userID):
    client = init()
    request: GetUserRequest = (
        GetUserRequest.builder()
        .user_id(userID)
        .user_id_type("open_id")
        .department_id_type("open_department_id")
        .build()
    )

    option = lark.RequestOption.builder().user_access_token(user_Access_Token).build()
    response: GetUserResponse = client.contact.v3.user.get(request, option)
    if not response.success():
        lark.logger.error(
            f"client.contact.v3.user.get failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}"
        )
        return

    return response


def GetTasksOfListRequest(page_token=None):
    client = init()
    request = (
        TasksTasklistRequest.builder()
        .tasklist_guid(tasklist_Guid)
        .page_size(100)
        .user_id_type("open_id")
        .build()
    )

    if page_token:
        request.page_token(page_token)

    option = lark.RequestOption.builder().user_access_token(user_Access_Token).build()
    response = client.task.v2.tasklist.tasks(request, option)
    if not response.success():
        lark.logger.error(
            f"client.task.v2.tasklist.tasks failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}"
        )
        return

    return response


def GetSubTaskOfListRequest(task_id, page_token=None):
    client = init()
    request: ListTaskSubtaskRequest = (
        ListTaskSubtaskRequest.builder()
        .task_guid(task_id)
        .page_size(50)
        .user_id_type("open_id")
        .build()
    )

    if page_token:
        request.page_token(page_token)

    option = lark.RequestOption.builder().user_access_token(user_Access_Token).build()
    response: ListTaskSubtaskResponse = client.task.v2.task_subtask.list(
        request, option
    )
    if not response.success():
        lark.logger.error(
            f"client.task.v2.task_subtask.list failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}"
        )
        return

    return response


def GetCustomSectionRequest(section_Guid):
    client = init()
    request: GetSectionRequest = (
        GetSectionRequest.builder()
        .section_guid(section_Guid)
        .user_id_type("open_id")
        .build()
    )
    option = lark.RequestOption.builder().user_access_token(user_Access_Token).build()
    response: GetSectionResponse = client.task.v2.section.get(request, option)
    if not response.success():
        # The subtask may not have a title / 子任务可能没有标题
        # lark.logger.error(
        #     f"client.task.v2.section.get failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
        return

    return response


def GetNameByUserID(userID):
    response_data = json.loads(
        lark.JSON.marshal(GetUserNameRequest(userID).data, indent=4)
    )
    try:
        name = response_data["user"]["name"]
        return name
    except KeyError as e:
          return None


def GetMemberNameByLoop(members):
    members_name = []
    if CheckExists(members):
        for member in members:
            if member.role == "assignee":
                members_name.append(GetNameByUserID(member.id))

    return "、".join(members_name)


def GetSectionNameBySectionGuid(section_guid):
    response = GetCustomSectionRequest(section_guid)
    if response:
        section_data = response.data
        return section_data.section.name

    return None


def GetSectionGuid(taskList):
    if taskList:
        for i in taskList:
            return i.section_guid

    return None


def GetCustomPlannedWorkingHoursFields(custom_fields):
    """本项目自定义字段：计划工时"""
    if custom_fields is None:
        return None

    for field in custom_fields:
        if field.name == "计划工时":
            return field.number_value
    return None


def GetCustomDevelopmentHoursFields(custom_fields):
    """本项目自定义字段：开发工时"""
    if custom_fields is None:
        return None

    for field in custom_fields:
        if field.name == "开发工时":
            return field.number_value
    return None


def CheckExcelFilePath(executable_dir):
    if getattr(sys, "frozen", False):
        file_path = executable_dir
    else:
        file_path = os.getcwd()
    return os.path.join(file_path, "output.xlsx")


def TimeChange(unixTime):
    if unixTime:
        timestamp = int(unixTime) / 1000
        if (
            datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")
            == "1970-01-01 08:00:00"
        ):
            return None
        else:
            return datetime.fromtimestamp(timestamp)
    else:
        return None


def CheckExists(unCheckItem):
    return unCheckItem != None or unCheckItem != ""


if __name__ == "__main__":
    main()
