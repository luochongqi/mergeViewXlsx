#!/usr/bin/env python
# -*- coding: UTF-8 -*-
# @Time    : 2023/6/16 10:04 //这里时创建该文件的时间
# @Author  : Luo hao ye  //这里写自己的名字
# @File    : formalV0.5.py  //文件名
# @ProjectName: mergeViewXlsx //项目名称
# @Software: PyCharm //IDE
import xlwings as xw
import time
import sys

# 打开excel程序，默认设置：程序可见，只打开不新建工作簿，屏幕更新关闭
app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = False

# 获取sheet1
wb = app.books.open('05.xlsx')
sht1 = xw.sheets.active

# 获取工作表有多少行数据
rng = sht1.range('A1').expand('table')  # 以第一列为基础，直至遇到第一个空单元格，获取工作表有多少行数据
sht_rows = rng.rows.count - 1  # 需要排除第一行表头


# 主函数
def main():
    # 视图的数量
    views = 8
    # 获得物料业务数据维护的数量
    tr_rows = get_rows(views)
    # 获得分组
    list_group = grouping(tr_rows)
    # 维护视图数据复制到基本视图
    copy_to_complete(list_group, tr_rows)


# 提示消息的输出函数
def message(statement):
    print(statement)


# 程序正常退出函数
def procedure_exit(statement):
    message(statement)
    time.sleep(3)
    app.quit()
    sys.exit()


# 根据views值获得物料业务数据数量的函数
def get_rows(views):
    tr_rows = 0
    if sht_rows % views == 0:
        tr_rows = int(sht_rows / views)
        message(f"此次维护的物料数据数量为：{tr_rows} 条")
    else:
        procedure_exit(f"错误！不符合维护视图导出报表的规则，请检查源文件，程序即将自动退出！")
    return tr_rows


# 分组函数
def grouping(tr_rows):
    first_index = 1
    last_index = sht_rows
    return list(range(first_index, last_index, tr_rows))


# 复制函数
def copy_to_complete(list_group, tr_rows):
    i = 1
    while i + 1 <= len(list_group):  # 第一层循环，视图切换
        fix_index = 2
        index = list_group[i] + 1
        for j in range(1, tr_rows + 1):  # 第二层循环，维护视图数据复制到基本视图
            use_fix_index = fix_index + j - 1
            use_index = index + j - 1
            # print(use_index, use_fix_index)
            source_row = f'A{use_index}:K{use_index}'
            dest_row = f'A{use_fix_index}:K{use_fix_index}'
            sht1.range(source_row).copy(sht1.range(dest_row))
        i = i + 1
    # 删除1 + tr_rows + 1~sht_rows + 1之间的所有行
    for row in range(2 + tr_rows, sht_rows + 2):
        sht1.range('A' + str(2 + tr_rows)).api.EntireRow.Delete()


# 调用主函数
main()
# sht.range('A17:J17').copy(sht.range('A2:J2'))
# sht.range('A1:J1').api.EntireRow.Delete()

# 保存
wb.save('05-物料业务数据维护报表.xlsx')

# 退出excel
app.quit()
