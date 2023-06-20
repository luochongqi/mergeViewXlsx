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

app = None
wb = None
sht1 = None
sht2 = None


# 主函数
def main():
    filename = '临时.xlsx'
    open_file(filename)
    deal_file()
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


# 处理xlsx文件
def deal_file():
    a = sht1.range('A1:A2').value
    print(a)


# 关闭xlsx文件
def close_file(f_name):
    global app
    global wb

    # 保存
    wb.save('(已处理)' + f_name)

    # 退出excel
    app.quit()


main()
