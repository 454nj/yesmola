# -*- encoding: utf-8 -*-
import time
import requests
import json
from fake_useragent import UserAgent
import pandas as pd
from pyecharts import options as opts
from pyecharts.charts import Bar, Line, Pie, Grid, Map, Timeline


class EpidemicSpider(object):
    def __init__(self):
        # self.sjc = int(time.time()*1000)
        # self.url = f'https://view.inews.qq.com/g2/getOnsInfo?name=disease_h5&callback=_&_={self.sjc}'
        self.url = 'https://api.inews.qq.com/newsqa/v1/query/inner/publish/modules/list?modules=statisGradeCityDetail,diseaseh5Shelf'
        self.ua = UserAgent()
        self.headers = {
            'user_agent': self.ua.random,
        }

    def parse_response(self):
        response = requests.post(self.url, headers=self.headers).text
        data_th = json.loads(response)['data']['diseaseh5Shelf']['areaTree'][0]['children']  # 拿到json数据中 children 的值
        # 设置一个数据集
        data_set = []
        for i in data_th:
            # 创建空字典
            data_dict = {}
            # 地区名称
            data_dict['province'] = i['name']
            # 疫情数据
            data_dict['nowConfirm'] = i['total']['nowConfirm']  # 现有确诊
            data_dict['confirm'] = i['total']['confirm']  # 累计确诊
            data_dict['addconfirm'] = i['today']['confirm']  # 新增确诊
            data_dict['dead'] = i['total']['dead']  # 死亡人数
            data_dict['heal'] = i['total']['heal']  # 治愈人数
            data_dict['wzz'] = i['total']['wzz']  # 本土无症状确诊
            data_dict['provinceLocalConfirm'] = i['total']['provinceLocalConfirm']  # 地区本地确诊
            data_set.append(data_dict)
        fd = pd.DataFrame(data_set)  # 制成表格
        # fd.to_excel('全国地区疫情数据.xls')  # 保存Excel
        fd2 = fd.sort_values(by=['nowConfirm'], ascending=False)  # 根据 'nowConfirm' 以多到少向下排序
        self.drawing(fd)

    def drawing(self, fd):
        """-----------------------------绘制饼状图图表---------------------------------"""
        pie = (
            Pie()
            .add(
                '',
                [list(i) for i in zip(fd['province'].values.tolist(), fd['heal'].values.tolist())],
                radius=['40%', '80%']  # 改变内圆与外圆
            )
            .set_global_opts(
                legend_opts=opts.LegendOpts(orient='vertical', pos_top='70%', pos_left='70%')
            )
            .set_series_opts(label_opts=opts.LabelOpts(formatter='{b}:{c}'))
        )
        # 生成一个前端页面
        # pie.render('pie.html')

        """---------------------------绘制折线图---------------------------------------------"""
        line = (
            Line()
            .add_xaxis(list(fd['province'].values))
            .add_yaxis('治愈人数', fd['heal'].values.tolist())
            .add_yaxis('死亡人数', fd['dead'].values.tolist())
            .set_global_opts(
                title_opts=opts.TitleOpts(title='死亡人数与治愈人数')
            )
        )
        # line.render('line.html')

        """---------------------------绘制柱形图---------------------------------------------"""

        bar = Bar()
        bar.add_xaxis(list(fd['province'].values.tolist()))
        bar.add_yaxis('治愈人数', fd['heal'].values.tolist())
        bar.add_yaxis('新增确诊', fd['addconfirm'].values.tolist())
        bar.set_global_opts(
            title_opts=opts.TitleOpts(title='各地区疫情数据'),
            datazoom_opts=[opts.DataZoomOpts()]  # 伸缩条
        )
        # bar.render('bar.html')

        """---------------------------绘制地图---------------------------------------------"""
        china_map = (
            Map()
            # china--意为中国地图
            .add('现有确诊', [list(i) for i in zip(fd['province'], fd['nowConfirm'])], 'china', is_map_symbol_show=False)
            .add('治愈人数', [list(i) for i in zip(fd['province'], fd['heal'])], 'china', is_map_symbol_show=False)
            .add('累计确诊', [list(i) for i in zip(fd['province'], fd['confirm'])], 'china', is_map_symbol_show=False)
            .add('新增确诊', [list(i) for i in zip(fd['province'], fd['addconfirm'])], 'china', is_map_symbol_show=False)
            .set_global_opts(
                title_opts=opts.TitleOpts(title='各地区疫情情况'),
                visualmap_opts=opts.VisualMapOpts(max_=500, is_piecewise=True),
                legend_opts=opts.LegendOpts(pos_left='90%', pos_top='60%')
            )
        )
        china_map.render('map.html')

        """---------------------------在一个页面放上面全部图形---------------------------------------------"""
        # grid = (
        #     Grid(init_opts=opts.InitOpts(width='4000px', height='1500px'))
        #     .add(pie, grid_opts=opts.GridOpts(pos_top='50%', pos_right='50%'))
        #     .add(line, grid_opts=opts.GridOpts(pos_bottom='50%', pos_right='50%'))
        #     # .add(bar, grid_opts=opts.GridOpts(pos_bottom='50%', pos_left='50%'))
        #     .add(pie, grid_opts=opts.GridOpts(pos_top='50%', pos_left='50%'))
        # )
        # grid.render('grid.html')


if __name__ == '__main__':
    e = EpidemicSpider()
    e.parse_response()
