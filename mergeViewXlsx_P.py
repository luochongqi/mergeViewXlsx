#!/usr/bin/env python
# -*- coding: UTF-8 -*-
# @Time    : 2023/6/29 13:44 //这里时创建该文件的时间
# @Author  : Luo HaoYe  //这里写自己的名字
# @File    : mergeViewXlsx_P.py  //文件名
# @ProjectName: mergeViewXlsx //项目名称
# @Software: PyCharm //IDE
import xlwings as xw
import time
import sys
import os
import threading
import traceback
from tqdm import tqdm
import keyboard
import pywintypes
from win32com.client import DispatchEx
from xlwings._xlwindows import COMRetryObjectWrapper

# 全局宏
APP_ERROR = 1
NORMAL_ERROR = 0

# 全局数据，全部为空
app = None
wb = None
sht1 = None  # sheet0，源数据sheet
sht2 = None  # sheet1，工作区间sheet
filename = ''

# 输出文字样式变量
head_color_font_green = '\033[1;32;40m'  # 绿色高亮
head_color_font_red = '\033[1;33;40m'  # 红色高亮
tail_color_font = '\033[0m'

# 各视图字段在excel中列的索引号
i_quality_view = []
i_procure_view = []
i_sale_view = []
i_storage_view = []
i_workplan_view = []
i_mrp_view = []
i_basic_view = []


# 主函数
def main():
    try:
        # 获取视图索引模板文件
        get_template()
        # 源sheet的最终数据的末行计数变量
        data_row = 2
        # 打开文件，获得处理数据的总行数
        sht1_rows = open_file()
        # 检查正在处理的数据是否都为归档流程的数据
        check_gd(sht1_rows)
        # 计时器
        start_time = time.perf_counter()
        # 视图的数量
        views = 8
        # 获得流程列表、流程详细信息字典
        flows_list = sht1.range(f'BQ2:BQ{sht1_rows}').value
        fun_ret_distribute_flow = distribute_flow(flows_list, sht1_rows)
        tr_rows = get_rows(views, sht1_rows, 1)  # 获得物料业务数据维护的总数量
        # flow_list = fun_ret_distribute_flow[0]  流程列表
        flow_dic = fun_ret_distribute_flow[1]
        # 多流程遍历处理
        for flow_k, flow_v in flow_dic.items():
            # 从源sheet将单个流程数据复制到工作sheet中，并获得当前流程处理数据的行数
            sht2_rows = span_sheet_copy_go(flow_k, flow_v)
            # 获得当前流程物料业务数据维护的数量
            distribute_tr_rows = get_rows(views, sht2_rows, 0)
            # 获得分组
            list_group = grouping(distribute_tr_rows, sht2_rows)
            # 维护视图数据复制到基本视图
            copy_to_complete(list_group, distribute_tr_rows, sht2_rows)
            # 从工作sheet将已处理的数据复制回源sheet，并更新data_row
            data_row = span_sheet_copy_come(flow_k, distribute_tr_rows, data_row)
        # 其余数据处理
        remaining_data_process(tr_rows, sht1_rows, data_row)
        # 输出总用时
        end_time = time.perf_counter() - start_time
        print(f"\n程序总运行时间为：{end_time}s。")
        # 关闭文件
        close_file(filename)
        # 退出程序
        new_thread1 = threading.Thread(target=thread_exit_hand, name="T1")
        new_thread1.daemon = True  # 创建守护线程，当主线程执行完毕时，子线程不管有没有执行完都跟着结束
        new_thread1.start()
        time.sleep(5)
        sys.exit()
        # new_thread.join()
    except BaseException:
        traceback.print_exc()
        procedure_exit(f'遭遇到未预设的错误！！！', NORMAL_ERROR)


# 获取视图索引模板文件
def get_template():
    global i_quality_view
    global i_procure_view
    global i_sale_view
    global i_storage_view
    global i_workplan_view
    global i_mrp_view
    global i_basic_view
    # 初始化i_views，用于接收lists
    i_views = [i_quality_view, i_procure_view, i_sale_view, i_storage_view, i_workplan_view, i_mrp_view, i_basic_view]
    # 打开并读取template.ini文件
    template_name = 'Config\\template.txt'
    with open(template_name) as file_object:
        lines = file_object.readlines()
    i = 0
    for line in lines:
        if i != 7:
            line = line.strip()
            if line != '\n':
                i_views[i] = line.split(' ')
            else:
                i_views[i] = []
            i += 1
    # 为各视图添加字段索引
    if len(i_views) == 7:
        i_quality_view = i_views[0]
        i_procure_view = i_views[1]
        i_sale_view = i_views[2]
        i_storage_view = i_views[3]
        i_workplan_view = i_views[4]
        i_mrp_view = i_views[5]
        i_basic_view = i_views[6]
    else:
        procedure_exit(f'错误！template.txt文件出错！', NORMAL_ERROR)


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


# 自动退出函数
def thread_exit_hand():
    print("\n程序执行完毕，按回车键退出程序！（5秒后自动退出）")
    while True:
        if keyboard.is_pressed('enter'):
            os._exit(0)


# 提示消息的输出函数
def message(statement):
    print(head_color_font_red + statement + tail_color_font)


# 程序正常退出函数
def procedure_exit(statement, error):
    message(statement)
    time.sleep(5)
    if error == 1:
        app.kill()
    sys.exit()


# 打开文件函数
def open_file():
    # 声明全局变量
    global app
    global wb
    global sht1
    global sht2
    global filename

    # 打开excel程序，默认设置：程序可见，只打开不新建工作簿，屏幕更新关闭
    app = get_excel_app()
    if not app:
        procedure_exit(f"打开excel的程序遇到问题，程序即将退出！", APP_ERROR)
    app.display_alerts = False
    app.screen_updating = False

    # 获取sheet1
    print("导入文件要求如下：")
    print("1、文件为excel文件，后缀名为.xlsx；")
    print("2、excel文件必须和该程序处于同一个目录下，且需要输入完整的文件名（例如：xxx.xlsx；注意：不用带路径）；\n\n")
    while True:
        filename = input("请输入需要处理的完整文件名：").strip()
        f_split_list = filename.split('.')
        if len(f_split_list) < 2 or f_split_list[1] != 'xlsx':
            message(f"文件名后缀发生错误，请检查！")
            press_key = input(f"\n输入'q'退出程序！输入其他任意信息程序继续！\n")
            if press_key == 'q':
                procedure_exit(f'即将退出程序！', APP_ERROR)
            else:
                continue
        elif not os.path.exists(filename):
            message(f"该文件不存在！")
            press_key = input(f"\n输入'q'退出程序！输入其他任意信息程序继续！\n")
            if press_key == 'q':
                procedure_exit(f'即将退出程序！', APP_ERROR)
            else:
                continue
        else:
            break

    wb = app.books.open(filename)
    sht1 = xw.sheets.active
    sht2 = wb.sheets.add(name='sheet1', after=sht1)  # 新建sheet1，在sheet0后

    # 获取工作表有多少行数据
    rng = sht1.range('BP1').expand('table')  # 以AY列为基础，直至遇到第一个空单元格，获取工作表有多少行数据
    in_sht1_rows = rng.rows.count - 1  # 需要排除第一行表头
    return in_sht1_rows


# 检查是否皆为归档数据
def check_gd(in_sht1_rows):
    in_rng = sht1.range(f'BP2:BP{in_sht1_rows + 1}')
    in_list = in_rng.value
    for i in in_list:
        if i.find('归档') == -1:
            procedure_exit(f'错误！该批数据中存在非存档流程的数据，程序即将退出！', APP_ERROR)


# 其余数据处理
def remaining_data_process(in_tr_rows, in_sht1_rows, in_data_row):
    # 删除BP、BQ两项辅助列，采用重复删除BP已达到删除这两列的效果
    for row in range(1, 3):
        sht1.range('BP:BP').api.EntireColumn.Delete()
    # 重复删除A{data_row} '2 + in_tr_rows' ~ 'sht1_rows + 2'已达到删除冗余数据的效果
    for row in range(2 + in_tr_rows, in_sht1_rows + 2):
        sht1.range(f'A{in_data_row}').api.EntireRow.Delete()
    # 增加序列号
    serial_list = list(range(1, in_tr_rows + 1))
    sht1.api.Columns(1).Insert()
    sht1.range('A1').value = '序号'
    sht1.range('A2').options(transpose=True).value = serial_list


# 跨sheet区域复制——源sheet到工作sheet
def span_sheet_copy_go(in_flow_k, in_flow_v):
    print(f'\n{head_color_font_green}正在处理流程：{in_flow_k}{tail_color_font}\n')
    source = f'A{in_flow_v[0]}:BQ{in_flow_v[1]}'
    dest = f'A2'
    sht1.range(source).copy(sht2.range(dest))
    in_sht2_rows = in_flow_v[1] - in_flow_v[0] + 1
    return in_sht2_rows


# 跨sheet区域复制——工作sheet到源sheet
def span_sheet_copy_come(in_flow_k, distribute_tr_rows, in_data_row):
    source = f'A2:BQ{distribute_tr_rows + 2 - 1}'  # 例如流程只有一条物料数据，此时1+2-1才是正确的
    dest = f'A{in_data_row}'
    sht2.range(source).copy(sht1.range(dest))
    in_data_row = in_data_row + distribute_tr_rows
    print(f'\n{head_color_font_green}流程处理结束：{in_flow_k}{tail_color_font}\n')
    return in_data_row


# 分流程函数
def distribute_flow(in_flows_list, in_sht1_rows):
    div_value = 2  # 逻辑上的第0行数据，在磁盘的Excel上实际是第2行数据
    i = 0
    old = in_flows_list[i]  # 流程元素
    distribute_flow_list = [old]  # 记录所有流程元素的列表
    in_flow_dic = {old: [div_value]}
    for flow in in_flows_list:
        if flow != old:
            distribute_flow_list.append(flow)
            in_flow_dic.get(old).append(i + div_value - 1)  # 触发流程末行索引添加
            old = flow
            in_flow_dic[old] = [i + div_value]  # 触发流程首行索引添加
        i = i + 1
        if i + div_value == in_sht1_rows + 1:  # 最后一行不会触发流程末行索引添加代码，需要进行边界情况处理
            in_flow_dic.get(old).append(in_sht1_rows + 1)
    message(f'本次维护的物料业务维护视图流程为：{len(in_flow_dic)} 个')
    return distribute_flow_list, in_flow_dic


# 根据views值获得物料业务数据数量的函数
def get_rows(views, in_sht1_rows, flag):
    tr_rows = 0
    if in_sht1_rows % views == 0:
        tr_rows = int(in_sht1_rows / views)
        if flag == 1:
            message(f"此次维护的物料数据数量总共为：{tr_rows} 条")
        else:
            message(f"该流程维护的物料数据数量为：{tr_rows} 条")
    else:
        procedure_exit(f"错误！不符合维护视图导出报表的规则，程序即将退出！", APP_ERROR)
    return tr_rows


# 分组函数
def grouping(distribute_tr_rows, in_sht2_rows):
    first_index = 1
    last_index = in_sht2_rows
    list_group = list(range(first_index, last_index, distribute_tr_rows))
    # 边界情况处理：当只处理1条物料数据时，需要给分组额外加上[8]
    if distribute_tr_rows == 1:
        list_group.append(8)
        return list_group
    else:
        return list_group


# source_row = f'A{use_index}:K{use_index}'
# dest_row = f'A{use_fix_index}:K{use_fix_index}'
# sht.range(source_row).copy(sht.range(dest_row))

# 复制函数——核心函数
def copy_core(cell, use_index, use_fix_index):
    source = f'{cell}{use_index}'
    dest = f'{cell}{use_fix_index}'
    sht2.range(source).copy(sht2.range(dest))


# 视图复制
def copy_view(use_index, use_fix_index, i_views):
    for i in i_views:
        copy_core(i, use_index, use_fix_index)


# 复制函数分支流控制函数
def copy_branch_control(view_num, *args):
    # 财务视图：0    质量视图：1      采购视图：2      销售试图：3
    # 仓储视图：4    工作计划视图：5        MRP视图：6     基本视图：7
    if view_num == 1:
        copy_view(args[0], args[1], i_quality_view)
    elif view_num == 2:
        copy_view(args[0], args[1], i_procure_view)
    elif view_num == 3:
        copy_view(args[0], args[1], i_sale_view)
    elif view_num == 4:
        copy_view(args[0], args[1], i_storage_view)
    elif view_num == 5:
        copy_view(args[0], args[1], i_workplan_view)
    elif view_num == 6:
        copy_view(args[0], args[1], i_mrp_view)
    elif view_num == 7:
        copy_view(args[0], args[1], i_basic_view)


# 复制函数
def copy_to_complete(list_group, distribute_tr_rows, in_sht2_rows):
    print(f"\n复制任务进度：")
    c_items = range(1, len(list_group))  # 视图的索引范围在0~7
    for i in tqdm(c_items, desc='处理中', ncols=80):  # 第一层循环，视图切换
        fix_index = 2
        index = list_group[i] + 1
        for j in range(1, distribute_tr_rows + 1):  # 第二层循环，其余视图数据复制到财务视图
            use_fix_index = fix_index + j - 1
            use_index = index + j - 1
            # print(use_index, use_fix_index)
            copy_branch_control(i, use_index, use_fix_index)
    # 删除1 + distribute_tr_rows + 1~distribute_tr_rows + 1之间的所有行
    print(f"\n删除任务进度：")
    r_items = range(2 + distribute_tr_rows, in_sht2_rows + 2)
    for row in tqdm(r_items, desc='处理中', ncols=80):  # 只删除被顶在2 + distribute_tr_rows的这一行即可
        sht2.range('A' + str(2 + distribute_tr_rows)).api.EntireRow.Delete()


# 关闭文件函数
def close_file(f_name):
    # 使源sheet为active的
    sht1.activate()
    # 保存
    wb.save('(已处理)' + f_name)
    # 退出excel
    app.kill()


# 调用主函数
main()
# sht.range('A17:J17').copy(sht.range('A2:J2'))
# sht.range('A1:J1').api.EntireRow.Delete()
