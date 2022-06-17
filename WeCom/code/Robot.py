# coding: utf-8
# @Time    : 2022/5/19
# @Author  : caijinrong
# @Email   : kamwtsai@gmail.com
# @Link    : https://github.com/KamWTsai/AutoSendMsg

import requests
import base64
import hashlib

from requests_toolbelt.multipart.encoder import MultipartEncoder

class Robot:
    send_url = ""
    upload_url = ""

    def __init__(self, key):
        self.send_url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=" + key
        self.upload_url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key=" + key + "&type=file"

    def send_text(self, text):
        headers = {"Content-Type": "application/json"}
        data = {
            "msgtype": "text",
            "text": {
                "content": text
                # ,
                # "mentioned_list":["wangqing","@all"],
                # "mentioned_mobile_list":["13800001111","@all"]
            }
        }
        result = requests.post(self.send_url, headers=headers, json=data)
        return result

    def send_img(self, img_path):
        with open(img_path, 'rb') as file:  # 转换图片成base64格式
            data = file.read()
            b64encodestr = base64.b64encode(data)
            img_b64_data = str(b64encodestr, 'utf-8')

        with open(img_path, 'rb') as file:  # 图片的MD5值
            md = hashlib.md5()
            md.update(file.read())
            img_md5_data = md.hexdigest()

        headers = {"Content-Type": "application/json"}
        data = {
            "msgtype": "image",
            "image": {
                "base64": img_b64_data,
                "md5": img_md5_data
            }
        }
        result = requests.post(self.send_url, headers=headers, json=data)
        return result

    def upload_file(self, file_path):
        file_name = file_path.replace('\\','/').split('/')[-1]
        m = MultipartEncoder(
            fields={
                'filename': 'media',
                'file': (file_name, open(file_path, 'rb'), 'application/octet-stream')
            }
        )
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36',
            'Content-Type': 'multipart/form-data',
            'Content-Disposition': 'form-data; name="media";filename="test.txt"; filelength=20'
        }
        result = requests.post(self.upload_url, data=m, headers=headers).json()
        return result["media_id"]

    def send_file(self, file_path):
        for fp in file_path:
            media_id = self.upload_file(fp)
            headers = {"Content-Type": "application/json"}
            data = {
                "msgtype": "file",
                "file": {
                    "media_id": media_id
                }
            }
            result = requests.post(self.send_url, json=data, headers=headers)
