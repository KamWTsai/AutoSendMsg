# coding: utf-8
# @Time    : 2022/5/19
# @Author  : caijinrong
# @Email   : kamwtsai@gmail.com
# @Link    : https://github.com/KamWTsai/AutoSendMsg

import update_file
import excel_screen_shot
import send_mail
import schedule
import time
import os

current_path = os.path.abspath(__file__)
dir_path = os.path.dirname(os.path.dirname(current_path))

# 删除截图
def delete_pic(img_path = dir_path + os.sep + "pic" + os.sep):
    ls=os.listdir(img_path)
    for i in ls:
        rm_path=os.path.join(img_path,i)
        os.remove(rm_path)


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
    img_name_list = []
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

    # 发送邮件
    sm = send_mail.SendMail(settings, img_name_list[0])  # 暂时只支持发1张图片，后续再拓展
    sm.sendMail()

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

if __name__ == '__main__':
    settings = send_mail.loadSettings()

    is_regular = settings["是否定时"]
    if is_regular == 1:
        run_at_regular_time(settings)
    else:
        run_at_once(settings)


