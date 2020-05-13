#!/usr/bin/env python
# -*- coding: utf-8 -*-
# author: seven_lit 
# name:frank_hertz.py time: 2020/5/13

import xlrd
import xlwt

#  文件名
input_filename = 'frank_hertz.xlsx'
output_filename = 'frank_answers.xls'


def read_data():
    """
    读取原工作表的数据
    :return: 将读取结果储存在列表
    """
    data = xlrd.open_workbook(input_filename)
    sheet = data.sheet_by_index(0)
    peak_volt = sheet.row_values(rowx=3, start_colx=2, end_colx=8)
    valley_volt = sheet.row_values(rowx=4, start_colx=2, end_colx=8)
    return peak_volt, valley_volt


def step_subtraction(peak_volt, valley_volt):
    """
    逐差法计算
    :param peak_volt: 波峰电压
    :param valley_volt: 波谷电压
    :return: 逐差法的计算结果
    """
    output_peak = [0, 0, 0]
    output_valley = [0, 0, 0]
    for i in range(0, 3):
        output_peak[i] = (peak_volt[i + 3] - peak_volt[i]) / 3
    for i in range(3):
        output_valley[i] = (valley_volt[i + 3] - valley_volt[i]) / 3
    return output_peak, output_valley


def write_data(output_peak, output_valley):
    """
    创建并写入工作表计算的结果
    :param output_peak: 波峰逐差计算结果
    :param output_valley: 波谷逐差计算结果
    :return: xls结果工作表
    """
    out_sum = 0
    # 创建excel文件并命名工作表
    my_workbook = xlwt.Workbook()
    my_sheet = my_workbook.add_sheet('answers')
    #  设置单元格格式，水平、垂直居中
    style = xlwt.XFStyle()
    fmt = xlwt.Alignment()
    fmt.horz = 0x02
    fmt.vert = 0x01
    style.alignment = fmt
    # 设置三列的列宽
    a = my_sheet.col(2)
    b = my_sheet.col(3)
    c = my_sheet.col(4)
    a.width = 256 * 40
    b.width = 256 * 40
    c.width = 256 * 40
    # 标题打印
    my_sheet.write(2, 2, '（第4个波峰（谷）-第1个波峰（谷））/3.0', style)
    my_sheet.write(2, 3, '（第5个波峰（谷）-第2个波峰（谷））/3.0', style)
    my_sheet.write(2, 4, '（第6个波峰（谷）-第3个波峰（谷））/3.0', style)
    my_sheet.write(3, 1, '波峰', style)
    my_sheet.write(4, 1, '波谷', style)
    my_sheet.write(5, 1, '平均', style)
    # 计算结果打印
    for i in range(3):
        my_sheet.write(3, 2+i, output_peak[i], style)
        out_sum = out_sum + output_peak[i]
    for i in range(3):
        my_sheet.write(4, 2+i, output_valley[i], style)
        out_sum = out_sum + output_valley[i]
    out_data = out_sum / 6
    my_sheet.write_merge(5, 5, 2, 4, out_data, style)
    #  保存xls文件
    my_workbook.save(output_filename)


# 执行程序
peak, valley = read_data()
out_peak, out_valley = step_subtraction(peak, valley)
write_data(out_peak, out_valley)