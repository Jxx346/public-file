
一、岗位职责
1、店铺统筹与规划：负责公司1688店铺的整体运营规划，包括店铺定位、产品布局、视觉呈现及日常维护，提升店铺星级及各项综合指标。
2、产品上架与优化：负责口腔个护3C产品（如电动牙刷、冲牙器、牙齿美白仪等）的上架、标题优化、属性完善及详情页策划，提升搜索排名与转化率。
3、活动报名与推广：策划并报名1688平台各类活动（如伙拼、实力商家、商人节等），合理运用网销宝、标王等付费推广工具，控制ROI，提升流量与销量。
4、数据分析与复盘：每日监控店铺核心数据（流量、点击、转化、客单价等），定期进行竞品分析与市场调研，输出优化方案并落地执行。
5、供应链与库存协同：对接内部供应链及仓储部门，跟进3C产品的备货、发货及售后处理，确保履约时效，降低客诉率。
6、B2B客户维护：协助处理大客户询盘、分销商招募及老客复购维护，建立稳定的B端客户池。
二、任职要求
1、经验要求：1-3年1688国内站运营经验，有3C数码、小家电或口腔护理类目经验者优先。
2、平台技能：精通1688后台操作，熟悉平台规则、流量玩法及活动报名机制，有成功打造爆款或提升店铺层级的案例。
3、数据敏感度：具备较强的数据分析能力，能熟练使用生意参谋等工具，通过数据驱动业务决策。
4、沟通协作：具备良好的跨部门沟通能力，能高效对接美工、客服、仓储及供应链团队。
5、加分项：有口腔个护品牌方或3C代工厂1688运营经验；熟悉B2B分销体系搭建；自带优质供应链或客户资源。
三、我们的优势：
1，项目成熟 细分类目Top 1品牌，自带爆款流量
2，工厂管理和产品开发成熟，拥有多个私模爆款
四、岗位福利：
1、【工作环境】：5星级写字楼，宽敞明亮，提供舒适的办公环境，距离地铁口杨美站A出口50米，上下班方便。
2、【工作时间】：上午9:00-12:00，下午14:00-18:30，午休2小时，大小周。
3、【薪酬体系】：底薪+绩效，入职购买五险一金；
4、【假期福利】：享受国家规定的各类假期，下午茶福利&生日会，各种团建活动，年会。







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
    
69abe2d5559dea0015c7725d

![Uploading image.png…]()

def main(args):
    a = ['抖音', '月儿弯弯佳润口腔护理专卖店', '2026-07-02', "zanwushuj", 6]
    data = post_kai(a)
