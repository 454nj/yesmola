# -*- encoding: utf-8 -*-
from requests_html import HTMLSession
import sys
import time
import os, xlwt, xlrd
from xlutils.copy import copy
# 构造请求对象
session = HTMLSession()


class QqGroup(object):
    def __init__(self):
        self.user_input_qqId = input('请输入QQ群号:')
        self.url = "https://qun.qq.com/cgi-bin/qun_mgr/search_group_members"
        self.headers = {
            'authority': 'qun.qq.com',
            'accept': 'application/json, text/javascript, */*; q=0.01',
            'accept-language': 'en,zh-CN;q=0.9,zh;q=0.8,en-US;q=0.7',
            'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'cookie': 'pgv_pvid=291657330; _tc_unionid=b784c004-6aaf-43dc-8051-af5b7d0a2d07; tvfe_boss_uuid=2b9d98e281e913f0; eas_sid=j1E6L4U2g438h2u606Z4Q2y0J8; pac_uid=0_9b63a76cb5722; RK=yLPtSpgR/e; ptcz=7a7e0933626057e75606a92c2a5c83abce2c9689b6344b0770a7382416ac8761; iip=0; fqm_pvqid=b69b3e9f-58c0-43ef-a1db-751236f872c8; ptui_loginuin=2357349368; _qpsvr_localtk=0.11578864995714144; uin=o2357349368; p_uin=o2357349368; traceid=9de0bd49b4; pgv_info=ssid=s3316945842; skey=@FkAsjKkhW; pt4_token=ZGj*JiiwrYXzmDIu46GLnIejVAzwa64b2p1qMfTW5vw_; p_skey=ODHvJNuutA2g7j-J6MagxigzFYpnInBk*asyDdB4JEw_',
            'origin': 'https://qun.qq.com',
            'referer': 'https://qun.qq.com/member.html',
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="101", "Google Chrome";v="101"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.0.0 Safari/537.36',
            'x-requested-with': 'XMLHttpRequest'
        }

    def parse_group(self):
        try:
            start = 0
            end = 20
            while True:
                payload = f"gc={self.user_input_qqId}&st={start}&end={end}&sort=0&bkn=1901121513"
                response = session.post(self.url, headers=self.headers, data=payload).json()['mems']  # 通过键取值
                self.parse_friends_list(response)  # 将值传递给下一个函数解析
                start = end + 1
                end += 21
        except Exception as e:
            print(e)  # 打印异常
            sys.exit(0)  # 退出进程

    def parse_friends_list(self, response):
        for res in response:
            name = res['nick'].replace('&nbsp;', ' ')  # 昵称
            group_name = res['card'].replace('&nbsp;', ' ')  # 群昵称
            qqNumber = res['uin']  # qq号
            age = res['qage']  # Q龄
            gender = res['g']  # 性别
            # 加入群组时间
            join_time = res['join_time']
            time_array = time.localtime(join_time)
            jtime = time.strftime("%Y-%m-%d %H:%M:%S", time_array)  # 时间戳转换
            # 最后说话时间
            last_speak_time = res['last_speak_time']
            time_array = time.localtime(last_speak_time)
            ltime = time.strftime("%Y-%m-%d %H:%M:%S", time_array)  # 时间戳转换
            print(name)
            print(group_name)
            print(qqNumber)
            print(age)
            print(gender)
            print(jtime)
            print(ltime)
            print('===' * 30)
            data = {
                f'{self.user_input_qqId}': [name, group_name, qqNumber, age, gender, jtime, ltime]
            }
            self.save(data)

    def save(self, data):
        # 获取表的名称
        sheet_name = [i for i in data.keys()][0]
        # 创建保存excel表格的文件夹
        # os.getcwd() 获取当前文件路径
        os_mkdir_path = os.getcwd() + '/qq群数据/'
        # 判断这个路径是否存在，不存在就创建
        if not os.path.exists(os_mkdir_path):
            os.mkdir(os_mkdir_path)
        # 判断excel表格是否存在           工作簿文件名称
        os_excel_path = os_mkdir_path + '数据.xls'
        if not os.path.exists(os_excel_path):
            # 不存在，创建工作簿(也就是创建excel表格)
            workbook = xlwt.Workbook(encoding='utf-8')
            """工作簿中创建新的sheet表"""  # 设置表名
            worksheet1 = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
            """设置sheet表的表头"""
            sheet1_headers = ('昵称', '群昵称', 'qq号', 'Q龄', '性别（-1:未知, 0:男, 1:女）', '加入群组时间', '最后说话时间')
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
            excel_headers_tuple = ('昵称', '群昵称', 'qq号', 'Q龄', '性别（-1:未知, 0:男, 1:女）', '加入群组时间', '最后说话时间')
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
    q = QqGroup()
    q.parse_group()


