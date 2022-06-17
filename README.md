# AutoSendMsg

[TOC]

## 1. 总概

本项目通过Python实现【自动更新Excel数据并发送文件】，分别可用于发送邮件、发送至企业微信群聊

- 项目链接：https://github.com/KamWTsai/AutoSendMsg

- 博客链接：https://blog.csdn.net/u013068739/article/details/125334586

**项目原理**：

1. 大数据平台调度作业每日跑出数据，将数据导出至MySQL
2. 通过Python自动化实现
   - Excel通过Power Query插件从MySQL拉取更新数据
   - 检测Excel是否更新成功
   - 截图发送邮件或发送至企业微信群聊

> 注意：本项目只实现以上【2】中的功能，【1】作为前置条件必须满足

**Python代码逻辑步骤**：

1. 读取设置`settings.json`
2. 更新Excel文件
3. 检测Excel文件是否更新成功
4. 截图
5. 发送



## 2. 邮件(Mail)

#### 2.1 项目结构

```
Mail
│  settings.json
│  sendMailLog.log
│  text.txt
│  需要更新发送的Excel文件.xlsx
│
├─code
│      Run.py
│      update_file.py
│      excel_screen_shot.py
│      send_mail.py
│
└─pic
```

`Mail/`目录下：

- settings.json：设置

- sendMailLog.log：发送日志
- text.txt：可设置邮件正文、图片
- 需要更新发送的Excel文件.xlsx

`Mail/code/`目录下：

- Run.py：运行核心主要逻辑
- update_file.py：更新文件并检测文件是否更新成功
- excel_screen_shot.py：截图
- send_mail.py：发送邮件

#### 2.2 生成邮箱第三方客户端密码

根据不同邮箱生成

#### 2.3 设置说明

```json
{
	"服务器地址": "smtp.xx.com",
	"端口": 465,
	"发件人名称": "xxx",
	"发件人地址": "xxx@xx.com",
	"密码": "xxx",
	"收件人地址": [
		"xxx@xx.com",
		"xxx@xx.com"
	],
	"标题": "xxx",
	"正文路径": "text.txt",
	"文件路径": [
		"需要更新发送的Excel文件.xlsx"
	],
	"截图sheet": "xx",
	"截图区域": ["A1:S30"],
	"时间": "10:00",
	"是否定时": 0,
	"是否更新": 1,
	"更新检测": {
		"check": 1,
		"sleep_time": 1800,
		"sheet": "检测",
		"cell": "B1"
	}
}
```

1. 服务器地址：根据实际需要修改
2. 端口：根据实际需要修改
3. 发件人名称
4. 发件人地址
5. 密码：创建的客户端密码
6. 收件人地址
7. 标题：邮件标题
8. 正文路径：邮件正文的text路径
9. 文件路径：要发送的文件路径
10. 截图sheet：填写要截图的Sheet的名称。也可以用数字，下标从0开始，0是Sheet1
11. 截图区域：按照Excel的单元格区域，中间的冒号为**英文符号**
12. 时间：定时发送的时间，24小时制，中间的冒号为**英文符号**。在【是否定时】置为1时生效
13. 是否定时：1为是，0为否（即为**立即发送**）
14. 是否更新：1为是，0为否。（因为存在自己手动更新后，只需要发送的情况）
15. 更新检测：
    - check：是否检测更新成功，1为是，0为否
    - sleep_time：检测后若未更新成功，等待x秒后重新更新再检测
    - sheet：检测页面的Sheet名称
    - cell：检测的单元格，现只能检测单元格数值是否为1，不为1则认为**更新失败**

#### 2.4 直接运行

在`Mail/code/`目录下，运行cmd

```shell
python Run.py
```



## 3. 企微机器人(WeCom)

#### 3.1 项目结构

```
WeCom
│  settings.json
│  sendMsgLog.log
│  需要更新发送的Excel文件.xlsx
│
├─code
│      Run.py
│      update_file.py
│      excel_screen_shot.py
│      Robot.py
│
└─pic
```

`WeCom/`目录下：

- settings.json：设置

- sendMsgLog.log：发送日志
- 需要更新发送的Excel文件.xlsx

`WeCom/code/`目录下：

- Run.py：运行核心主要逻辑
- update_file.py：更新文件并检测文件是否更新成功
- excel_screen_shot.py：截图
- Robot.py：发送至企业微信群聊

#### 3.2 设置说明

```json
{
	"测试Key": "xxx",
	"正式Key": "xxx",
	"文字": "%d月%d日数据",
	"文件路径": [
		"需要更新发送的Excel文件.xlsx"
	],
	"截图sheet": "xx",
	"截图区域" : ["A1:S30"],
	"时间": "10:00",
	"是否定时": 0,
	"是否测试发送": 1,
	"是否正式发送": 0,
	"是否发送文字": 1,
	"是否发送截图": 1,
	"是否发送文件": 1,
	"是否更新": 1,
	"更新检测": {
		"check": 1,
		"sleep_time": 1800,
		"sheet": "检测",
		"cell": "B1"
	}
}
```

1. 测试key：企业微信用于测试的群的机器人key
2. 正式key：企业微信正式发送的群的机器人key
3. 文字：如需发送特定文字可自定义修改
4. 文件路径：要发送的文件路径
5. 截图sheet：填写要截图的Sheet的名称。也可以用数字，下标从0开始，0是Sheet1
6. 截图区域：按照Excel的单元格区域，中间的冒号为**英文符号**
7. 时间：定时发送的时间，24小时制，中间的冒号为**英文符号**。在【是否定时】置为1时生效
8. 是否定时：1为是，0为否（即为**立即发送**）
9. 是否测试发送：是否发送到测试群，1为是，0为否
10. 是否正式发送：是否发送到正式群，1为是，0为否
11. 是否发送文字：1为是，0为否
12. 是否发送截图：1为是，0为否
13. 是否发送文件：1为是，0为否
14. 是否更新：1为是，0为否。（因为存在自己手动更新后，只需要发送的情况）
15. 更新检测：
    - check：是否检测更新成功，1为是，0为否
    - sleep_time：检测后若未更新成功，等待x秒后重新更新再检测
    - sheet：检测页面的Sheet名称
    - cell：检测的单元格，现只能检测单元格数值是否为1，不为1则认为**更新失败**

#### 3.3 直接运行

在`WeCom/code/`目录下，运行cmd

```shell
python Run.py
```



## 4. 说明/脚本原理/注意事项

1. **原理**：Python在**后台**打开指定路径的Excel并更新截图（Excel在任务栏不可见，仅可在任务管理器中看见）

2. **调用接口**：Python调用的微软Excel接口不稳定，容易报错，解决方案如下：

   - 程序运行前，请确保**完全关闭Excel**（包括后台进程）。在cmd中输入以下命令：

     ```
     taskkill /F /IM EXCEL.EXE
     ```

   - 若仍未能解决请重启电脑

3. **更新**：Python通过第三方xlwings库调用微软Excel接口，相当于在Excel中手动点击【数据】→【全部刷新】

   ```python
   wb.api.RefreshAll()  # 调用微软Excel接口
   ```

4. **检测**：刷新文件后，读取“检测”Sheet中指定单元格（比如B1），判断检测是否成功

5. **截图**：调用微软给出的接口，相当于以下过程：选中Excel中指定截图区域单元格→按住键盘上的Ctrl+C→复制到系统中的剪切板→粘贴到邮件/发送到群聊中

6. **截图实现原理**：调用微软Excel接口，将指定区域复制为图片粘贴到剪切板中。因此在邮件自动更新机器人运行过程中，**请勿复制Ctrl+C其他内容**（包括但不限于文本、图片、文件等）

   > 1. 否则Python脚本会将当前你复制的内容替换Excel复制的图片，然后将你复制的内容粘贴到邮件中发送出去
   > 2. **请勿同时运行多个邮件自动发送机器人**，因为也是共用一个系统剪切板，可能A机器人复制的图片会粘贴到B机器人负责的邮件里面去发送

## 5. 其他

#### 5.1 定时运行Python脚本

##### 5.1.1 通过windows计划任务（推荐）

右键此电脑→管理→系统工具→任务计划程序→任务计划程序库→Microsoft→Windows→创建任务

引用博客仅供参考：[win10设置Python程序定时运行(设置计划任务)](https://www.cnblogs.com/JesseP/p/10816192.html)

##### 5.1.2 通过项目中的schedule第三方库

1. 在settings.json中，可以设置【是否定时】为1，再设置【时间】
2. 运行Run.py即可定时

> 注意：但容易因为报错导致程序停止，需要重启定时