# coding: utf-8
# @Time    : 2022/5/19
# @Author  : caijinrong
# @Email   : kamwtsai@gmail.com
# @Link    : https://github.com/KamWTsai/AutoSendMsg

import Robot
import excel_screen_shot
import update_file
import time, datetime
import schedule
import json
import os

current_path = os.path.abspath(__file__)
dir_path = os.path.dirname(os.path.dirname(current_path))

# 机器人发送
def robot_send(robot, settings, img_name_list):
    file_path = [dir_path + os.sep + settings["文件路径"][0]]

    if settings["是否发送文字"] == 1:
        now = datetime.datetime.now()
        robot.send_text(settings["文字"] %(now.month, now.day))
    if settings["是否发送文件"] == 1:
        robot.send_file(file_path)
    if settings["是否发送截图"] == 1:
        for img_name in img_name_list:
            robot.send_img(img_name)

 # 写入日志
def write_to_log(content):
    log_text = '[' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + '] ' + content
    print(log_text)
    with open(dir_path + os.sep + "sendMsgLog.log", 'a', encoding='utf-8') as f:
        f.writelines(log_text +'\n')

# 删除截图
def delete_pic(img_path = dir_path + os.sep + "pic" + os.sep):
    ls=os.listdir(img_path)
    for i in ls:
        rm_path=os.path.join(img_path,i)
        os.remove(rm_path)

# 立即发送
def run_at_once(settings):
    print("[开始运行]")

    file_path = dir_path + os.sep + settings["文件路径"][0]

    # 更新
    if settings["是否更新"] == 1:
        update_attempts = 0
        update_flag = False
        while not update_flag:  # 更新成功
            try:
                update_flag = update_file.update_file(file_path, settings)
                if update_flag == False:
                    print("[等待更新] " + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + " 已尝试【%d】次,【%d】s后再更新…" % (
                        update_attempts+1, settings["更新检测"]["sleep_time"]))
                    time.sleep(settings["更新检测"]["sleep_time"])
            except Exception as e:
                print(e)
            finally:
                update_attempts += 1
            if update_attempts == 3 and not update_flag:
                print("[失败]更新3次仍无法成功")
                return

    # 给excel截图
    sheet = settings["截图sheet"]
    range = settings["截图区域"]
    screenshot_attempts = 0
    screenshot_flag = False
    while not screenshot_flag or len(img_name_list)==0:  # 截图成功
        try:
            img_name_list = excel_screen_shot.excel_save_img(file_path, sheet, range, save_path = dir_path + os.sep + "pic" + os.sep)
            screenshot_flag = True
        except Exception as e:
            print(e)
        finally:
            screenshot_attempts += 1
        if screenshot_attempts == 3 and not screenshot_flag:
            print("[失败]截图3次仍无法成功")
            return

    # 测试发送
    if settings["是否测试发送"] == 1:
        robot = Robot.Robot(settings["测试Key"])
        robot_send(robot, settings, img_name_list)
        write_to_log("测试")

    # 正式发送
    if settings["是否正式发送"] == 1:
        robot = Robot.Robot(settings["正式Key"])
        robot_send(robot, settings, img_name_list)
        write_to_log("正式")

    # 删除截图
    delete_pic()

# 定时功能
def run_at_regular_time(settings):
    regular_time = settings['时间']
    print("[定时发送] " + regular_time)
    # 定时每天的这个时间发送
    schedule.every().day.at(regular_time).do(run_at_once, settings)
    while True:
        schedule.run_pending()
        time.sleep(1)

def loadSettings():
    with open(dir_path + os.sep + "settings.json", 'r', encoding='utf-8') as f:
        settings = json.load(f)
    return settings


if __name__ == '__main__':
    # 加载设置
    settings = loadSettings()

    is_regular = settings["是否定时"]
    if is_regular == 1:
        run_at_regular_time(settings)
    else:
        run_at_once(settings)
