# coding: utf-8
# @Time    : 2022/5/19
# @Author  : caijinrong
# @Email   : kamwtsai@gmail.com
# @Link    : https://github.com/KamWTsai/AutoSendMsg

import xlwings as xw
import time


# 更新文件，且默认检测
def update_file(file_path, settings):
    flag = False  # 判断更新成功的标志

    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(file_path)

        # 判断是否已经更新，已经更新过的就不用再更新了
        if settings["更新检测"]["check"]  == 1:
            sht = wb.sheets(settings["更新检测"]["sheet"])
            cell_value = sht.range(settings["更新检测"]["cell"]).value
            if cell_value == 1:
                return True

        # 更新
        print("[更新文件中] " + file_path)
        wb.api.RefreshAll()

        attempts = 0  # 每30s检测一次
        # 90s都没更新成功，先保存退出，等待下一次更新
        while attempts < 3 and not flag:  # 未更新成功，重试
            print("第%d轮循环"%(attempts+1))
            time.sleep(30)
            attempts += 1
            if settings["更新检测"]["check"]  == 1:
                # 获取判断是否更新的cell
                sht = wb.sheets(settings["更新检测"]["sheet"])
                cell_value = sht.range(settings["更新检测"]["cell"]).value
                print("[检测值] " + str(cell_value))
                if cell_value == 1:
                    flag = True
    except Exception as e:
        print(e)
    finally:
        wb.save()
        wb.close()
        app.quit()
        # try:
        #     app.kill()
        # except Exception as e:
        #     print(e)

    check_str = ""
    if settings["更新检测"]["check"] == 1:
        if flag == True:
            check_str = "：成功"
        else:
            check_str = "：失败"
    else:
        flag = True  # 不检测的默认更新成功
    print("[更新结束%s]" % check_str)
    return flag