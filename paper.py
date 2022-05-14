# -*- encoding: utf-8 -*-
from requests_html import HTMLSession
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
import os, xlwt, xlrd
from xlutils.copy import copy
from urllib.parse import quote, unquote
# 构造请求对象
session = HTMLSession()


class ZhiWangPaper(object):
    def __init__(self):
        self.user_input = input('请输入关键词：')
        self.keyword_quote = quote(self.user_input)
        self.SearchSql = input('请输入关键词的SearchSql：')
        self.url = 'https://kns.cnki.net/kns8/Brief/GetGridTableHtml'
        self.ua = UserAgent()
        self.headers = {
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'Cookie': 'Ecp_notFirstLogin=gpT11i; Ecp_ClientId=2220315225011999024; cnkiUserKey=694288d7-5a50-b6db-8a96-f326aa098ec2; knsLeftGroupSelectItem=1%3B2%3B; SID_sug=126004; Ecp_session=1; Ecp_loginuserbk=gz0367; _pk_ref=%5B%22%22%2C%22%22%2C1651044073%2C%22https%3A%2F%2Fwww.baidu.com%2Flink%3Furl%3DSKKVWRPXmmRPNGHAjXM62jDssl4sFfo8u5Cs0wBdXQa%26wd%3D%26eqid%3Db5dcfcb50000c5d3000000066268cfaa%22%5D; _pk_ses=*; LID=WEEvREcwSlJHSldSdmVpa3VEenplbUdORWhpV1ZtR0hUc3FTdXNoeG9kRT0=$9A4hF_YAuvQ5obgVAqNKPCYcEjKensW4IQMovwHtwkF4VYPoHbKxJwu0021u0021; Ecp_LoginStuts={"IsAutoLogin":false,"UserName":"gz0367","ShowName":"%E4%B8%9C%E8%8E%9E%E8%81%8C%E4%B8%9A%E6%8A%80%E6%9C%AF%E5%AD%A6%E9%99%A2","UserType":"bk","BUserName":"","BShowName":"","BUserType":"","r":"gpT11i"}; ASP.NET_SessionId=ig41rkp3fmaqr0icfmp0u55f; SID_kns8=25123162; _pk_id=67f2482a-fe12-4a7e-96d2-a0d1eaa41454.1647355805.3.1651044115.1651044073.; SID_kns_new=kns123162; CurrSortField=%e7%9b%b8%e5%85%b3%e5%ba%a6%2frelevant%2c(%e5%8f%91%e8%a1%a8%e6%97%b6%e9%97%b4%2c%27time%27)+desc; CurrSortFieldType=desc; Ecp_ClientIp=119.141.85.33; SID_docpre=006007; c_m_LinID=LinID=WEEvREcwSlJHSldSdmVpa3VEenplbUdORWhpV1ZtR0hUc3FTdXNoeG9kRT0=$9A4hF_YAuvQ5obgVAqNKPCYcEjKensW4IQMovwHtwkF4VYPoHbKxJwu0021u0021&ot=04%2f27%2f2022%2015%3a43%3a24; c_m_expire=2022-04-27%2015%3a43%3a24; dblang=ch; CurrSortField=%e7%9b%b8%e5%85%b3%e5%ba%a6%2frelevant%2c(%e5%8f%91%e8%a1%a8%e6%97%b6%e9%97%b4%2c%27time%27)+desc; CurrSortFieldType=desc',
            'Origin': 'https://kns.cnki.net',
            'Referer': 'https://kns.cnki.net/kns8/defaultresult/index',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.60 Safari/537.36'
        }

    def parse_response(self):
        print('当前下载第1页')
        first_payload = f"IsSearch=true&QueryJson=%7B%22Platform%22%3A%22%22%2C%22DBCode%22%3A%22CFLS%22%2C%22KuaKuCode%22%3A%22CJFQ%2CCDMD%2CCIPD%2CCCND%2CCISD%2CSNAD%2CBDZK%2CCCJD%2CCCVD%2CCJFN%22%2C%22QNode%22%3A%7B%22QGroup%22%3A%5B%7B%22Key%22%3A%22Subject%22%2C%22Title%22%3A%22%22%2C%22Logic%22%3A1%2C%22Items%22%3A%5B%7B%22Title%22%3A%22%E4%B8%BB%E9%A2%98%22%2C%22Name%22%3A%22SU%22%2C%22Value%22%3A%22{self.keyword_quote}%22%2C%22Operate%22%3A%22%25%3D%22%2C%22BlurType%22%3A%22%22%7D%5D%2C%22ChildItems%22%3A%5B%5D%7D%5D%7D%7D&PageName=defaultresult&DBCode=CFLS&KuaKuCodes=CJFQ%2CCDMD%2CCIPD%2CCCND%2CCISD%2CSNAD%2CBDZK%2CCCJD%2CCCVD%2CCJFN&CurPage=1&RecordsCntPerPage=20&CurDisplayMode=listmode&CurrSortField=&CurrSortFieldType=desc&IsSentenceSearch=false&Subject="
        first_response = session.post(self.url, headers=self.headers, data=first_payload).text
        # print(first_response)
        self.parse_data(first_response)
        for page in range(2, 21):
            print(f'当前下载第{page}页')
            other_payload = f"IsSearch=false&QueryJson=%7B%22Platform%22%3A%22%22%2C%22DBCode%22%3A%22CFLS%22%2C%22KuaKuCode%22%3A%22CJFQ%2CCDMD%2CCIPD%2CCCND%2CCISD%2CSNAD%2CBDZK%2CCCJD%2CCCVD%2CCJFN%22%2C%22QNode%22%3A%7B%22QGroup%22%3A%5B%7B%22Key%22%3A%22Subject%22%2C%22Title%22%3A%22%22%2C%22Logic%22%3A1%2C%22Items%22%3A%5B%7B%22Title%22%3A%22%E4%B8%BB%E9%A2%98%22%2C%22Name%22%3A%22SU%22%2C%22Value%22%3A%22{self.keyword_quote}%22%2C%22Operate%22%3A%22%25%3D%22%2C%22BlurType%22%3A%22%22%7D%5D%2C%22ChildItems%22%3A%5B%5D%7D%5D%7D%7D&SearchSql={self.SearchSql}&PageName=defaultresult&HandlerId=15&DBCode=CFLS&KuaKuCodes=CJFQ%2CCDMD%2CCIPD%2CCCND%2CCISD%2CSNAD%2CBDZK%2CCCJD%2CCCVD%2CCJFN&CurPage={page}&RecordsCntPerPage=20&CurDisplayMode=listmode&CurrSortField=&CurrSortFieldType=desc&IsSortSearch=false&IsSentenceSearch=false&Subject="
            first_response = session.post(self.url, headers=self.headers, data=other_payload).text
            # print(first_response)
            self.parse_data(first_response)

    def parse_data(self, first_response):
        soup = BeautifulSoup(first_response, 'lxml')
        tr_list = soup.find_all('tr')
        for tr in tr_list:
            td = BeautifulSoup(str(tr), 'lxml')
            titles = td.select('html body tr td.name a')  # 标题
            authors = td.select('html body tr td.author')  # 作者
            sources = td.select('html body tr td.source a')  # 来源
            datetimes = td.select('html body tr td.date')  # 发表时间
            databases = td.select('html body tr td.data')  # 数据库
            downloads = td.select('html body tr td.download a')  # 下载量
            for i, j, k, l, m, n in zip(titles, authors, sources, datetimes, databases, downloads):
                title = i.get_text()
                author = j.get_text().replace('\n', '')
                source = k.get_text()
                datetime = l.get_text().lstrip()
                database = m.get_text().lstrip()
                download = n.get_text()
                print(title)
                print(author)
                print(source)
                print(datetime)
                print(database)
                print(download)
                print('===' * 30)
                data = {
                    f'{self.user_input}': [title, author, source, datetime, database, download]
                }
                self.save(data)

    def save(self, data):
        # 获取表的名称
        sheet_name = [i for i in data.keys()][0]
        # 创建保存excel表格的文件夹
        # os.getcwd() 获取当前文件路径
        os_mkdir_path = os.getcwd() + '/知网论文数据/'
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
            sheet1_headers = ('标题', '作者', '来源', '发表时间', '数据库', '下载量')
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
            excel_headers_tuple = ('标题', '作者', '来源', '发表时间', '数据库', '下载量')
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
    z = ZhiWangPaper()
    z.parse_response()