#!/usr/bin/env python
# -*- coding: UTF-8 -*-
# @Time    : 2023/6/27 18:52 //这里时创建该文件的时间
# @Author  : Luo HaoYe  //这里写自己的名字
# @File    : test_basic.py  //文件名
# @ProjectName: mergeViewXlsx //项目名称
# @Software: PyCharm //IDE
import xlwings as xw
import sys
import pywintypes
from win32com.client import DispatchEx
from xlwings._xlwindows import COMRetryObjectWrapper
from pypinyin import pinyin, Style

app = None
wb = None
sht1 = None
sht2 = None
head_color_font_green = '\033[1;32;40m'
tail_color_font_green = '\033[0m'


# 主函数
def main():
    filename = '副本临时.xlsx'
    open_file(filename)
    check_gd(32)
    close_file(filename)


# 获得正确运行的excel程序
def get_excel_app():
    try:
        in_app = xw.App(visible=False, add_book=False)
        return in_app
    except pywintypes.com_error:
        try:
            _xl = COMRetryObjectWrapper(DispatchEx("ket.Application"))
            impl = xw._xlwindows.App(visible=False, add_book=False, xl=_xl)
            in_app = xw.App(visible=False, add_book=False, impl=impl)
            return in_app
        except pywintypes.com_error:
            return None


# 打开xlsx文件
def open_file(filename):
    global app
    global wb
    global sht1
    global sht2

    app = get_excel_app()
    if not app:
        sys.exit()
    app.display_alerts = False
    app.screen_updating = False

    wb = app.books.open(filename)
    sht1 = xw.sheets.active
    sht2 = wb.sheets.add(name='sheet1', after=sht1)


#  跨sheet区域复制
def span_sheet_copy(in_flow_k, in_flow_v):
    print(f'\n{head_color_font_green}正在处理流程：{in_flow_k}{tail_color_font_green}\n')
    source = f'A{in_flow_v[0]}:BQ{in_flow_v[1]}'
    dest = f'A2'
    sht1.range(source).copy(sht2.range(dest))
    in_sht2_rows = in_flow_v[1] - in_flow_v[0] + 1
    print(in_sht2_rows)


# 分流程函数
def distribute_flow(flows_list, sht_rows):
    div_value = 2  # 逻辑上的第0行数据，在磁盘的Excel上实际是第2行数据
    i = 0
    old = flows_list[i]  # 流程元素
    distribute_flow_list = [old]  # 记录所有流程元素的列表
    flow_dic = {old: [div_value]}
    for flow in flows_list:
        if flow != old:
            distribute_flow_list.append(flow)
            flow_dic.get(old).append(i + div_value - 1)  # 触发流程末行索引添加
            old = flow
            flow_dic[old] = [i + div_value]  # 触发流程首行索引添加
        i = i + 1
        print(i)
        if i + div_value == sht_rows:  # 最后一行不会触发流程末行索引添加代码，需要进行边界情况处理
            flow_dic.get(old).append(sht_rows)
    return distribute_flow_list, flow_dic


# 检测函数
def check_gd(in_sht1_rows):
    in_rng = sht1.range(f'BP2:BP{in_sht1_rows + 1}')
    in_list = in_rng.value
    for i in in_list:
        if i.find('归档') == -1:
            print(f'该批数据中存在非存档流程的数据，程序即将退出！')


# 处理xlsx文件
def deal_file():
    rng1 = sht1.range('BP1').expand('table')
    rows = rng1.rows.count - 1
    flows_list = sht1.range(f'BQ2:BQ{rows}').value
    fun_ret_distribute_flow = distribute_flow(flows_list, 33)
    # flow_list = fun_ret_distribute_flow[0]  流程列表
    flow_dic = fun_ret_distribute_flow[1]
    # 多流程遍历处理
    for flow_k, flow_v in flow_dic.items():
        span_sheet_copy(flow_k, flow_v)
    print(flows_list)
    print(flow_dic)
    # r_list = list(range(1, 13))
    # sht1.api.Columns(1).Insert()
    # sht1.range('A2').options(transpose=True).value = r_list


# 关闭xlsx文件
def close_file(f_name):
    global app
    global wb

    # 保存
    wb.save('(已处理)' + f_name)

    # 退出excel
    app.kill()


main()
