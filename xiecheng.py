# -*- encoding: utf-8 -*-
import requests
import json
import os, xlwt, xlrd
from xlutils.copy import copy
"""
城市ID{上海: 2, 北京: 1, 广州: 152, 深圳: 26}
"""


class XcSpider(object):
    def __init__(self):
        self.CityId = input('请输入城市ID：')
        self.url = "https://m.ctrip.com/restapi/soa2/18254/json/getAttractionList?_fxpcqlniredt=09031047314318028828&x-traceID=09031047314318028828-1646054807738-9064633"
        self.headers = {
            'content-type': 'application/json',
            'origin': 'https://m.ctrip.com',
            'referer': 'https://m.ctrip.com/webapp/you/gspoi/sight/2.html?seo=0&allianceid=4897&sid=155952&isHideNavBar=YES&from=https%3A%2F%2Fm.ctrip.com%2Fwebapp%2Fyou%2Fgsdestination%2Fplace%2F2.html%3Fseo%3D0%26ishideheader%3Dtrue%26secondwakeup%3Dtrue%26dpclickjump%3Dtrue%26allianceid%3D4897%26sid%3D155952%26ouid%3Dindex%26from%3Dhttps%253A%252F%252Fm.ctrip.com%252Fhtml5%252F',
            'accept-language': 'zh-CN,zh;q=0.9',
            'cookie': 'ibulanguage=CN; ibulocale=zh_cn; cookiePricesDisplayed=CNY; _gcl_au=1.1.2001712708.1646054591; _RF1=223.104.63.214; _RGUID=0731b0f7-45b5-4666-9828-888744fb269f; _RSG=cPKj5TFinS0VQo.4T8YeW9; _RDG=2868710522b1702c43085468305d1ce8b8; _bfaStatusPVSend=1; MKT_CKID=1646054594542.yi12k.1t3u; MKT_CKID_LMT=1646054594543; _ga=GA1.2.333705235.1646054595; _gid=GA1.2.1046662294.1646054595; appFloatCnt=2; nfes_isSupportWebP=1; GUID=09031047314318028828; nfes_isSupportWebP=1; MKT_Pagesource=H5; _bfs=1.4; _jzqco=%7C%7C%7C%7C1646054602232%7C1.1650478479.1646054594536.1646054655182.1646054672431.1646054655182.1646054672431.0.0.0.4.4; __zpspc=9.2.1646054672.1646054672.1%232%7Cwww.baidu.com%7C%7C%7C%25E6%2590%25BA%25E7%25A8%258B%7C%23; _bfi=p1%3D100101991%26p2%3D100101991%26v1%3D5%26v2%3D4; _bfaStatus=success; mktDpLinkSource=ullink; librauuid=MTPpuP1M6AmQCSUc; ibu_h5_lang=en; ibu_h5_local=en-us; _pd=%7B%22r%22%3A12%2C%22d%22%3A259%2C%22_d%22%3A247%2C%22p%22%3A260%2C%22_p%22%3A1%2C%22o%22%3A263%2C%22_o%22%3A3%2C%22s%22%3A263%2C%22_s%22%3A0%7D; Union=OUID=&AllianceID=4897&SID=155952&SourceID=&AppID=&OpenID=&exmktID=&createtime=1646054807&Expires=1646659606764; MKT_OrderClick=ASID=4897155952&AID=4897&CSID=155952&OUID=&CT=1646054806768&CURL=https%3A%2F%2Fm.ctrip.com%2Fwebapp%2Fyou%2Fgspoi%2Fsight%2F2.html%3Fseo%3D0%26allianceid%3D4897%26sid%3D155952%26isHideNavBar%3DYES%26from%3Dhttps%253A%252F%252Fm.ctrip.com%252Fwebapp%252Fyou%252Fgsdestination%252Fplace%252F2.html%253Fseo%253D0%2526ishideheader%253Dtrue%2526secondwakeup%253Dtrue%2526dpclickjump%253Dtrue%2526allianceid%253D4897%2526sid%253D155952%2526ouid%253Dindex%2526from%253Dhttps%25253A%25252F%25252Fm.ctrip.com%25252Fhtml5%25252F&VAL={"h5_vid":"1646054589723.2rr0y3"}; _bfa=1.1646054589723.2rr0y3.1.1646054589723.1646054806818.1.10.214062'
        }

    def get_data(self):
        for i in range(1, 101):
            print('当前下载第{}页'.format(i))
            payload = json.dumps({
                "index": f'{i}',
                "count": 20,
                "sortType": 1,
                "isShowAggregation": True,
                "districtId": self.CityId,  # 城市ID
                "scene": "DISTRICT",
                "pageId": "214062",
                "traceId": "f33070fa-82a6-6d22-2d18-164f0af07734",
                "extension": [
                    {
                        "name": "osVersion",
                        "value": "10.3.1"
                    },
                    {
                        "name": "deviceType",
                        "value": "ios"
                    }
                ],
                "filter": {
                    "filterItems": []
                },
                "crnVersion": "2020-09-01 22:00:45",
                "isInitialState": True,
                "head": {
                    "cid": "09031047314318028828",
                    "ctok": "",
                    "cver": "1.0",
                    "lang": "01",
                    "sid": "8888",
                    "syscode": "09",
                    "auth": "",
                    "xsid": "",
                    "extension": []
                }
            })
            response = requests.post(self.url, headers=self.headers, data=payload).json()
            self.parse(response)

    def parse(self, response):
        result_list = response['attractionList']
        for result in result_list:
            city = result['card']['districtName']  # 城市
            place = result['card']['poiName']  # 景区
            status = result['card']['openStatus']  # 状态
            score = result['card']['commentScore']  # 评分
            tickets = result['card']['priceTypeDesc']  # 门票
            distance = result['card']['distanceStr']  # 距离市中心
            url = result['card']['detailUrl']  # 详情链接
            print(city)
            print(place)
            print(status)
            print(score)
            print(tickets)
            print(distance)
            print(url)
            print('===' * 30)
            data = {
                f'{self.CityId}': [city, place, score, tickets, distance, url]
            }
            # self.save(data)

    def save(self, data):
        # 获取表的名称
        sheet_name = [i for i in data.keys()][0]
        # 创建保存excel表格的文件夹
        # os.getcwd() 获取当前文件路径
        os_mkdir_path = os.getcwd() + '/携程数据/'
        # 判断这个路径是否存在，不存在就创建
        if not os.path.exists(os_mkdir_path):
            os.mkdir(os_mkdir_path)
        # 判断excel表格是否存在           工作簿文件名称
        os_excel_path = os_mkdir_path + '知网论文数据.csv'
        if not os.path.exists(os_excel_path):
            # 不存在，创建工作簿(也就是创建excel表格)
            workbook = xlwt.Workbook(encoding='utf-8')
            """工作簿中创建新的sheet表"""  # 设置表名
            worksheet1 = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
            """设置sheet表的表头"""
            sheet1_headers = ('城市', '景区', '评分', '门票', '距离市中心', '详情链接')
            # 将表头写入工作簿
            for header_num in range(0, len(sheet1_headers)):
                # 设置表格长度
                worksheet1.col(header_num).width = 2560 * 3
                # 写入表头        行,    列,           内容
                worksheet1.write(0, header_num, sheet1_headers[header_num])
            # 循环结束，代表表头写入完成，保存工作簿
            workbook.save(os_excel_path)
        """=============================已有工作簿添加新表==============================================="""
        # 打开工作薄
        workbook = xlrd.open_workbook(os_excel_path)
        # 获取工作薄中所有表的名称
        sheets_list = workbook.sheet_names()
        # 如果表名称：字典的key值不在工作簿的表名列表中
        if sheet_name not in sheets_list:
            # 复制先有工作簿对象
            work = copy(workbook)
            # 通过复制过来的工作簿对象，创建新表  -- 保留原有表结构
            sh = work.add_sheet(sheet_name)
            # 给新表设置表头
            excel_headers_tuple = ('城市', '景区', '评分', '门票', '距离市中心', '详情链接')
            for head_num in range(0, len(excel_headers_tuple)):
                sh.col(head_num).width = 2560 * 3
                #               行，列，  内容，            样式
                sh.write(0, head_num, excel_headers_tuple[head_num])
            work.save(os_excel_path)
        """========================================================================================="""
        # 判断工作簿是否存在
        if os.path.exists(os_excel_path):
            # 打开工作簿
            workbook = xlrd.open_workbook(os_excel_path)
            # 获取工作薄中所有表的个数
            sheets = workbook.sheet_names()
            for i in range(len(sheets)):
                for name in data.keys():
                    worksheet = workbook.sheet_by_name(sheets[i])
                    # 获取工作薄中所有表中的表名与数据名对比
                    if worksheet.name == name:
                        # 获取表中已存在的行数
                        rows_old = worksheet.nrows
                        # 将xlrd对象拷贝转化为xlwt对象
                        new_workbook = copy(workbook)
                        # 获取转化后的工作薄中的第i张表
                        new_worksheet = new_workbook.get_sheet(i)
                        for num in range(0, len(data[name])):
                            new_worksheet.write(rows_old, num, data[name][num])
                        new_workbook.save(os_excel_path)


if __name__ == '__main__':
    x = XcSpider()
    x.get_data()