import xlwings as xw
import time
import sys
import os
import threading
from tqdm import tqdm
import keyboard
import pywintypes
from win32com.client import DispatchEx
from xlwings._xlwindows import COMRetryObjectWrapper

# 全局数据，全部为空
app = None
wb = None
sht1 = None
sht_rows = 0
filename = ''

# 各视图字段在excel中列的索引号
i_quality_view = ['AA', 'AB']
i_procure_view = ['V', 'W', 'X', 'Y', 'Z']
i_sale_view = ['L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'AE']
i_storage_view = ['AC', 'AD', 'AO', 'AP']
i_workplan_view = ['AW', 'AX']
i_mrp_view = ['AF', 'AG', 'AH', 'AI', 'AJ', 'AM', 'AN', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV']
i_basic_view = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'AK', 'AL']


# 主函数
def main():
    while True:
        # 打开文件
        open_file()
        # 计时器
        start_time = time.perf_counter()
        # 视图的数量
        views = 8
        # 获得物料业务数据维护的数量
        tr_rows = get_rows(views)
        # 获得分组
        list_group = grouping(tr_rows)
        # 维护视图数据复制到基本视图
        copy_to_complete(list_group, tr_rows)
        # 删除AY列
        sht1.range('BP:BP').api.EntireColumn.Delete()
        # 输出总用时
        end_time = time.perf_counter() - start_time
        print(f"\n程序总运行时间为：{end_time}s。")
        # 关闭文件
        close_file(filename)
        # 输入Y或N连续执行程序
        flag = input("\n是否需要继续处理其他文件？默认退出程序。（y/n）\n")
        if flag.lower() == 'y':
            continue
        elif flag.lower() == 'n':
            break
        else:
            break
    # 退出程序
    new_thread1 = threading.Thread(target=thread_exit_hand, name="T1")
    new_thread1.daemon = True  # 创建守护线程，当主线程执行完毕时，子线程不管有没有执行完都跟着结束
    new_thread1.start()
    time.sleep(5)
    sys.exit()
    # new_thread.join()


# 获得正确运行的excel程序
def get_excel_app():
    try:
        app = xw.App(visible=False, add_book=False)
        return app
    except pywintypes.com_error:
        try:
            _xl = COMRetryObjectWrapper(DispatchEx("ket.Application"))
            impl = xw._xlwindows.App(visible=False, add_book=False, xl=_xl)
            app = xw.App(visible=False, add_book=False, impl=impl)
            return app
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
    print(statement)


# 程序正常退出函数
def procedure_exit(statement):
    message(statement)
    time.sleep(3)
    app.quit()
    sys.exit()


# 打开文件函数
def open_file():
    # 声明全局变量
    global app
    global wb
    global sht1
    global sht_rows
    global filename

    # 打开excel程序，默认设置：程序可见，只打开不新建工作簿，屏幕更新关闭
    app = get_excel_app()
    if not app:
        procedure_exit(f"打开excel的程序遇到问题，程序即将退出！")
    app.display_alerts = False
    app.screen_updating = False

    # 获取sheet1
    print("导入文件要求如下：")
    print("1、文件为excel文件，后缀名为.xlsx；")
    print("2、每一个excel文件只允许包含一个‘物料业务数据维护视图申请’流程的数据；")
    print("3、excel文件必须和该程序处于同一个目录下，且需要输入完整的文件名（例如：xxx.xlsx；注意：不用带路径）；\n\n")
    while True:
        filename = input("请输入需要处理的完整文件名：").strip()
        f_split_list = filename.split('.')
        if len(f_split_list) < 2 or f_split_list[1] != 'xlsx':
            print(f"文件名后缀发生错误，请检查！")
            press_key = input(f"\n输入'q'退出程序！输入其他任意信息程序继续！\n")
            if press_key == 'q':
                procedure_exit(f'即将退出程序！')
            else:
                continue
        elif not os.path.exists(filename):
            print(f"该文件不存在！")
            press_key = input(f"\n输入'q'退出程序！输入其他任意信息程序继续！\n")
            if press_key == 'q':
                procedure_exit(f'即将退出程序！')
            else:
                continue
        else:
            break

    wb = app.books.open(filename)
    sht = xw.sheets.active

    # 获取工作表有多少行数据
    rng = sht.range('BP1').expand('table')  # 以BP列为基础，直至遇到第一个空单元格，获取当前sheet中的数据range
    sht_rows = rng.rows.count - 1  # 需要排除第一行表头


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
    list_group = list(range(first_index, last_index, tr_rows))
    # 边界情况处理：当只处理1条物料数据时，需要给分组额外加上[8]
    if tr_rows == 1:
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
    sht1.range(source).copy(sht1.range(dest))


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
def copy_to_complete(list_group, tr_rows):
    print(f"\n复制任务进度：")
    c_items = range(1, len(list_group))  # 视图的索引范围在0~7
    for i in tqdm(c_items, desc='处理中', ncols=80):  # 第一层循环，视图切换
        fix_index = 2
        index = list_group[i] + 1
        for j in range(1, tr_rows + 1):  # 第二层循环，其余视图数据复制到财务视图
            use_fix_index = fix_index + j - 1
            use_index = index + j - 1
            # print(use_index, use_fix_index)
            copy_branch_control(i, use_index, use_fix_index)
    # 删除1 + tr_rows + 1~sht_rows + 1之间的所有行
    print(f"\n删除任务进度：")
    r_items = range(2 + tr_rows, sht_rows + 2)
    for row in tqdm(r_items, desc='处理中', ncols=80):  # 只删除被顶在2 + tr_rows的这一行即可
        sht1.range('A' + str(2 + tr_rows)).api.EntireRow.Delete()


# 关闭文件函数
def close_file(f_name):
    global app
    global wb

    # 保存
    wb.save('(已处理)' + f_name)

    # 退出excel
    app.quit()


# 调用主函数
main()
# sht.range('A17:J17').copy(sht.range('A2:J2'))
# sht.range('A1:J1').api.EntireRow.Delete()
