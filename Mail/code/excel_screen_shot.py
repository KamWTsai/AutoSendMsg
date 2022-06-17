# coding: utf-8
# @Time    : 2022/5/19
# @Author  : caijinrong
# @Email   : kamwtsai@gmail.com
# @Link    : https://github.com/KamWTsai/AutoSendMsg

import xlwings as xw
from PIL import ImageGrab, Image
import time
import os

current_path = os.path.abspath(__file__)
dir_path = os.path.dirname(os.path.dirname(current_path))


# 截图，并返回图片名称的list
def excel_save_img(file_path, sheet, range, save_path=dir_path+os.sep+"pic"+os.sep):
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(file_path)
    sht = wb.sheets[sheet]

    img_name_list = []

    try:
        for r in range:
            range_val = sht.range(r)
            range_val.api.CopyPicture()
            time.sleep(1)  # 避免下一行获取代码过快而获取不到
            sht.api.Paste()
            pic = sht.pictures[0]  # 当前图片
            pic.api.Copy()  # 复制图片
            time.sleep(1)  # 避免下一行获取代码过快而获取不到
            img = ImageGrab.grabclipboard()  # 获取剪贴板的图片数据
            img = white_background(img)  # 给PNG无颜色单元格加白色背景
            img_name = file_path.split(os.sep)[-1].split('.')[0] + time.strftime("%m%d%H%M%S", time.localtime()) + ".png"
            img.save(save_path + img_name)  # 保存图片
            pic.delete()  # 删除sheet上的图片
            img_name_list.append(os.path.abspath(save_path + img_name))
            time.sleep(1)
            # print(img_name_list)
    except Exception as e:
        print(e)
    finally:
        wb.close()  # 不保存，直接关闭
        app.quit()  # 退出
        # app.kill()  # 强制杀掉这个进程
        pass

    return img_name_list

# 给PNG无颜色单元格加白色背景
def white_background(img):
    img = img.convert('RGBA')
    background = Image.new('RGB', img.size, (255, 255, 255))
    x, y = img.size
    background.paste(img, (0, 0, x, y), img)
    img = background
    return img