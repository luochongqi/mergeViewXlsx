#!/usr/bin/env python
# -*- coding: UTF-8 -*-
# @Time    : 2023/7/18 14:39 //这里时创建该文件的时间
# @Author  : Nana Xing  //这里写自己的名字
# @File    : test_fileread.py  //文件名
# @ProjectName: mergeViewXlsx //项目名称
# @Software: PyCharm //IDE
import os


i_quality_view = []
i_procure_view = []
i_sale_view = []
i_storage_view = []
i_workplan_view = []
i_mrp_view = []
i_basic_view = []

i_views = [i_quality_view, i_procure_view, i_sale_view, i_storage_view, i_workplan_view, i_mrp_view, i_basic_view]

filename = 'Config\\template.txt'

with open(filename) as file_object:
    lines = file_object.readlines()

i = 0
for line in lines:
    if line != '\n':
        line = line.strip()
        i_views[i] = line.split(' ')
        i += 1

print(i_views)
print(lines)

