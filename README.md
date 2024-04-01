这是一个简单的可以导出飞书「任务」栏中的「清单任务」下面的任务为 Excel 的工具

本项目使用飞书官网提供的 [开放平台](https://open.feishu.cn/?lang=zh-CN) 接口完成

## 特点

- 跨平台易于使用
- 直接导出为 Excel 方便团队、企业数据收集与统计
- 可继续自定义开发加入新的功能

## 使用

条件：
1. Git
2. python 解释器
3. 创建好的 [飞书应用](https://open.feishu.cn/app/)

### 用户

1. 将项目拉取到本地

`git clone git@github.com:brealinxx/larkTaskExport.git`

2. 建立环境、安装依赖

```zsh
cd larkTaskExport
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

可以直接复制上述几个命令，也可以将 `.venv` 更换为自定义的虚拟环境名称，也可更改为使用指定的 python 解释器路径

3. 打包程序

`pyinstaller main.spec`

4. 根据 `.env.example` 创建 `.env` 文件

在 dist 目录下创建 `.env` 文件，里面模仿 `.env.example` 所需然后填写 [APP_ID,APP_SECRET](https://open.feishu.cn/app/),TASKLIST_GUID

> TASKLIST_GUID: 找到飞书项目中所需要读取的任务清单，然后右键点击分享，仅复制里面的 `guid=`后面的值
> 例如：
> URL: `https://applink.feishu.cn/client/todo/task_list?guid=97d28f-72ff5-4833-afa0-dfa123e418633`
> 所需要的值: `97d28f-72ff5-4833-afa0-dfa123e418633`

5. 启动程序

在 dist 目录下找到名为 `main` 可执行文件，双击启动后会自动跳转到飞书登陆网址，完成登陆后，**复制登陆后的 URL 地址** 然后粘贴到程序中回车确认等待即可

6. 完成

生成的 Excel 会在 dist 目录下出现

### DEV

上述步骤大致相同，只不过如果需要自己加功能测试则需要把 `.env` 文件放在 larkTaskExport 目录下，然后正常 `python main.py` 启动程序即可，Excel 文件也会在当前文件夹下生成

## TODO

`1.auto useraccesstoken(complete) 2.multithreading 2.hidden subtask(optional)`

## 贡献
 
欢迎提交 PR 来贡献代码，或通过 Issues 提出新功能建议或报告错误。

## LICENSE

MIT