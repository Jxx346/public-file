# public-file
今日一条用来传文件
# 使用提醒:
# 1. xbot包提供软件自动化、数据表格、Excel、日志、AI等功能
# 2. package包提供访问当前应用数据的功能，如获取元素、访问全局变量、获取资源文件等功能
# 3. 当此模块作为流程独立运行时执行main函数
# 4. 可视化流程中可以通过"调用模块"的指令使用此模块

import xbot
from xbot import print, sleep
from .import package
from .package import variables as glv
import requests


def post_kai(data):
    url = "http://192.168.0.51:8081/api/rpa/finance/invoice-records"
    heardes = {
        "Content-Type": "application/json",
        "Authorization": "Bearer jryt_f7858828_x6WwVWtHEDNS_3j6yU8CkGR8PGk2KF9A"
    }
    file_path = data.get("file","")
    data = {
        "account_id":data.get("ID"),
        "invoice_date":data.get("invoice_date"),
        "platform_invoice_count": data.get("platform_invoice_count"),
        "buyer_invoice_count": data.get("buyer_invoice_count"),
        "file":file_path,
        "remark":data.get("remark")
    }

    with open(file_path,"rb") as f:
        files = {
            "file":("invoice.csv",f, "text/csv")
        }
        resu = requests.post(url,json=data,headers = heardes,files = files)
    if resu.status_code == 200:
        print(f"更新成功;data = {data}")
    else:
        print(f"更新失败;data = {data};")
        print(f"返回：{resu.text}")
    


def main(args):
    a = ['抖音', '月儿弯弯佳润口腔护理专卖店', '2026-07-02', "zanwushuj", 6]
    data = post_kai(a)
