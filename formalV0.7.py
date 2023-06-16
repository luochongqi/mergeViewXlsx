#!/usr/bin/env python
# -*- coding: UTF-8 -*-
# @Time    : 2023/6/16 18:08 //这里时创建该文件的时间
# @Author  : Nana Xing  //这里写自己的名字
# @File    : formalV0.7.py  //文件名
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
wb = app.books.open('07.xlsx')
sht = xw.sheets.active

# 获取工作表有多少行数据
rng = sht.range('AY1').expand('table')  # 以AY列为基础，直至遇到第一个空单元格，获取工作表有多少行数据
sht_rows = rng.rows.count - 1  # 需要排除第一行表头

# 各视图字段在excel中列的索引号
i_quality_view = ['AA', 'AB']
i_procure_view = ['V', 'W', 'X', 'Y', 'Z']
i_sale_view = ['L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U']
i_storage_view = ['AC', 'AD', 'AE', 'AO', 'AP']
i_workplan_view = ['AW', 'AX']
i_mrp_view = ['AF', 'AG', 'AH', 'AI', 'AJ', 'AM', 'AN', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV']
i_basic_view = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'AK', 'AL']


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
    # 删除AY列
    sht.range('AY:AY').api.EntireColumn.Delete()


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


# source_row = f'A{use_index}:K{use_index}'
# dest_row = f'A{use_fix_index}:K{use_fix_index}'
# sht.range(source_row).copy(sht.range(dest_row))

# 复制函数——核心函数
def copy_core(cell, use_index, use_fix_index):
    source = f'{cell}{use_index}'
    dest = f'{cell}{use_fix_index}'
    sht.range(source).copy(sht.range(dest))


# 质量视图复制
def copy_view_quality(use_index, use_fix_index):
    for i in i_quality_view:
        copy_core(i, use_index, use_fix_index)


# 采购视图复制
def copy_view_procure(use_index, use_fix_index):
    for i in i_procure_view:
        copy_core(i, use_index, use_fix_index)


# 销售视图复制
def copy_view_sale(use_index, use_fix_index):
    for i in i_sale_view:
        copy_core(i, use_index, use_fix_index)


# 仓储视图复制
def copy_view_storage(use_index, use_fix_index):
    for i in i_storage_view:
        copy_core(i, use_index, use_fix_index)


# 工作计划视图复制
def copy_view_workplan(use_index, use_fix_index):
    for i in i_workplan_view:
        copy_core(i, use_index, use_fix_index)


# MRP视图复制
def copy_view_mrp(use_index, use_fix_index):
    for i in i_mrp_view:
        copy_core(i, use_index, use_fix_index)


# 基本视图复制
def copy_view_basic(use_index, use_fix_index):
    for i in i_basic_view:
        copy_core(i, use_index, use_fix_index)


# 复制函数分支流控制函数
def copy_branch_control(view_num, *args):
    # 财务视图：0    质量视图：1      采购视图：2      销售试图：3
    # 仓储视图：4    工作计划视图：5        MRP视图：6     基本视图：7
    if view_num == 1:
        copy_view_quality(args[0], args[1])
    elif view_num == 2:
        copy_view_procure(args[0], args[1])
    elif view_num == 3:
        copy_view_sale(args[0], args[1])
    elif view_num == 4:
        copy_view_storage(args[0], args[1])
    elif view_num == 5:
        copy_view_workplan(args[0], args[1])
    elif view_num == 6:
        copy_view_mrp(args[0], args[1])
    elif view_num == 7:
        copy_view_basic(args[0], args[1])


# 复制函数
def copy_to_complete(list_group, tr_rows):
    i = 1  # i的取值范围在0~7，分别对应着8个视图,但是逆序的
    while i + 1 <= len(list_group):  # 第一层循环，视图切换
        fix_index = 2
        index = list_group[i] + 1
        for j in range(1, tr_rows + 1):  # 第二层循环，其余视图数据复制到财务视图
            use_fix_index = fix_index + j - 1
            use_index = index + j - 1
            # print(use_index, use_fix_index)
            copy_branch_control(i, use_index, use_fix_index)
        i = i + 1
    # 删除1 + tr_rows + 1~sht_rows + 1之间的所有行
    for row in range(2 + tr_rows, sht_rows + 2):
        sht.range('A' + str(2 + tr_rows)).api.EntireRow.Delete()


# 调用主函数
main()
# sht.range('A17:J17').copy(sht.range('A2:J2'))
# sht.range('A1:J1').api.EntireRow.Delete()

# 保存
wb.save('07-物料业务数据维护报表.xlsx')

# 退出excel
app.quit()
