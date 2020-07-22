import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
from numpy import sqrt, abs
import subprocess

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method,InstrumentError
#from GeneralMethod.Report import Report

# 求电阻箱每一级的示数
def calc_r(num_R):
    list_r = []
    num_R = int(num_R * 10)
    cnt = 0
    while num_R:
        list_r.append(num_R % 10 * (10 ** (cnt - 1)))
        num_R = num_R // 10
        cnt = cnt + 1
    return list_r

class Potentiometer:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含
    report_data_keys = [
        "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18"
        "R_11_avg", "R_12_avg", "R_13_avg", "R_21_avg", "R_22_avg", "R_23_avg" 
        "E_x", "dt_R_11", "dt_R_12", "dt_R_21", "dt_R_22", "u_R_11", "u_R_12", "u_R_21", "u_R_22"
        "S", "dt_S", "u_S", "u_Ex_Ex", "u_Ex"
        "R_0", "U_x", "U_0"
        "R_x", "dt_R_0", "u_R_0", "dt_U_0", "u_U_0", "dt_U_x", "u_U_x", "u_Rx_Rx", "u_R_x"
        "final_1", "final_2" 
    ]

    PREVIEW_FILENAME = "Preview.pdf"  # 预习报告模板文件的名称
    DATA_SHEET_FILENAME = "data.xlsx"  # 数据填写表格的名称
    REPORT_TEMPLATE_FILENAME = "Potentiometer_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment1/1051Report.docx"  # 最后生成实验报告的相对路径

    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {} # 存放需要填入实验报告的
        print("1051 电位差计及其应用\n1. 实验预习\n2. 数据处理")
        while True:
            try:
                oper = input("请选择: ").strip()
            except EOFError:
                sys.exit(0)
            if oper != '1' and oper != '2':
                print("输入内容非法！请输入一个数字1或2")
            else:
                break
        if oper == '1':
            print("现在开始实验预习")
            print("正在打开预习报告......")
            Method.start_file(self.PREVIEW_FILENAME)
        elif oper == '2':
            print("现在开始数据处理")
            print("即将打开数据输入文件......")
            # 打开数据输入文件
            Method.start_file(self.DATA_SHEET_FILENAME)
            input("输入数据完成后请保存并关闭excel文件，然后按回车键继续")
            # 从excel中读取数据
            self.input_data("./"+self.DATA_SHEET_FILENAME) # './' is necessary when running this file, but should be removed if run main.py
            print("数据读入完毕，处理中......")
            # 计算物理量
            self.calc_data()
            # 计算不确定度
            self.calc_uncertainty()
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            Method.start_file(self.REPORT_OUTPUT_FILENAME)
            print("Done!")

    '''
    从excel表格中读取数据
    @param filename: 输入excel的文件名
    @return none
    '''
    def input_data(self, filename):
        ws = xlrd.open_workbook(filename).sheet_by_name('1051')
        # 实验一
        list_R_11 = []
        list_R_21 = []
        list_R_12 = []
        list_R_22 = []
        list_R_13 = []
        list_R_23 = []
        col = 1
        for row in range(3, 6):
             list_R_11.append(float(ws.cell_value(row, col)))
        col = 2
        for row in range(3, 6):
            list_R_21.append(float(ws.cell_value(row, col)))
        col = 3
        for row in range(3, 6):
            list_R_12.append(float(ws.cell_value(row, col)))
        col = 4
        for row in range(3, 6):
            list_R_22.append(float(ws.cell_value(row, col)))
        col = 5
        for row in range(3, 6):
            list_R_13.append(float(ws.cell_value(row, col)))
        col = 6
        for row in range(3, 6):
            list_R_23.append(float(ws.cell_value(row, col)))
        self.data['list_R_11'] = list_R_11
        self.data['list_R_12'] = list_R_12
        self.data['list_R_13'] = list_R_13
        self.data['list_R_21'] = list_R_21
        self.data['list_R_22'] = list_R_22
        self.data['list_R_23'] = list_R_23
        row = 1
        col = 1
        self.data['num_t'] = float(ws.cell_value(row, col))
        row = 1
        col = 3
        self.data['num_E_N'] = float(ws.cell_value(row, col))
        # 实验二
        row = 8
        col = 1
        self.data['num_R_0'] = float(ws.cell_value(row, col))
        col = 3
        self.data['num_U_0'] = float(ws.cell_value(row, col))
        col = 5
        self.data['num_U_x'] = float(ws.cell_value(row, col))
    
    '''
    进行数据处理

    '''
    def calc_data(self):
        # 实验一
        num_R_11_avg = Method.average(self.data['list_R_11'])
        num_R_12_avg = Method.average(self.data['list_R_12'])
        num_R_13_avg = Method.average(self.data['list_R_13'])
        num_R_21_avg = Method.average(self.data['list_R_21'])
        num_R_22_avg = Method.average(self.data['list_R_22'])
        num_R_23_avg = Method.average(self.data['list_R_23'])
        self.data['num_R_11_avg'] = num_R_11_avg
        self.data['num_R_12_avg'] = num_R_12_avg
        self.data['num_R_13_avg'] = num_R_13_avg
        self.data['num_R_21_avg'] = num_R_21_avg
        self.data['num_R_22_avg'] = num_R_22_avg
        self.data['num_R_23_avg'] = num_R_23_avg
        num_E_x = self.data['num_E_N'] / num_R_11_avg * num_R_12_avg
        self.data['num_E_x'] = num_E_x
        # 实验二
        num_R_x = self.data['num_U_x'] / self.data['num_U_0'] * self.data['num_R_0']
        self.data['num_R_x'] = num_R_x

    '''
    计算所有的不确定度
    '''

    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    def calc_uncertainty(self):
        # 实验一
        list_a = [5, 0.5, 0.2, 0.1, 0.1]
        list_r = calc_r(self.data['num_R_11_avg'])
        num_dt_R_11 = InstrumentError.resistance_box(list_a, list_r, 0.02)
        list_r = calc_r(self.data['num_R_21_avg'])
        num_dt_R_21 = InstrumentError.resistance_box(list_a, list_r, 0.02)
        list_r = calc_r(self.data['num_R_12_avg'])
        num_dt_R_12 = InstrumentError.resistance_box(list_a, list_r, 0.02)
        list_r = calc_r(self.data['num_R_22_avg'])
        num_dt_R_22 = InstrumentError.resistance_box(list_a, list_r, 0.02)
        num_u_R_11 = num_dt_R_11 / sqrt(3)
        num_u_R_21 = num_dt_R_21 / sqrt(3)
        num_u_R_12 = num_dt_R_12 / sqrt(3)
        num_u_R_22 = num_dt_R_22 / sqrt(3)
        self.data.update({"num_dt_R_11": num_dt_R_11, "num_dt_R_12": num_dt_R_12, "num_dt_R_21": num_dt_R_21, "num_dt_R_22": num_dt_R_22})
        self.data.update({"num_u_R_11": num_u_R_11, "num_u_R_12": num_u_R_12, "num_u_R_21": num_u_R_21, "num_u_R_22": num_u_R_22})
        # 仪器灵敏度及误差
        num_S = 9 / abs(self.data['num_R_22_avg'] - self.data['num_R_23_avg']) * 1e3
        num_dt_S = 0.2 / num_S
        num_u_S = num_dt_S / sqrt(3)
        self.data.update({"num_S": num_S, "num_dt_S": num_dt_S, "num_u_S": num_u_S})
        # 合成不确定度
        num_u_Ex_Ex = sqrt(((1 / self.data['num_R_11_avg'] - 1 / (self.data['num_R_11_avg'] + self.data['num_R_21_avg'])) ** 2) * (num_u_R_11 ** 2) + (num_u_R_21 / (self.data['num_R_11_avg'] + self.data['num_R_21_avg'])) ** 2 + ((1 / self.data['num_R_12_avg'] - 1 / (self.data['num_R_12_avg'] + self.data['num_R_22_avg'])) ** 2) * (num_u_R_12 ** 2) + (num_u_R_22 / (self.data['num_R_12_avg'] + self.data['num_R_22_avg'])) ** 2)
        num_u_Ex = self.data['num_E_x'] * num_u_Ex_Ex
        self.data.update({"num_u_Ex_Ex": num_u_Ex_Ex, "num_u_Ex": num_u_Ex_Ex})
        # 输出带不确定度的最终结果: 不确定度保留一位有效数字, 物理量结果与不确定度首位有效数字对齐
        self.data['final_1'] = Method.final_expression(self.data['num_E_x'], self.data['num_u_Ex'])
        print(self.data['final_1'])

        # 实验二
        # R_0不确定度
        list_a = [5, 0.5, 0.2, 0.1]
        list_r = calc_r(self.data['num_R_0'])
        num_dt_R_0 = InstrumentError.resistance_box(list_a, list_r, 0.02)
        num_u_R_0 = num_dt_R_0 / sqrt(3)
        self.data.update({"num_dt_R_0": num_dt_R_0, "num_u_R_0": num_u_R_0})
        # U_0不确定度
        num_dt_U_0 = InstrumentError.dc_potentiometer(0.01, self.data['num_U_0'], 0.1)
        num_u_U_0 = num_dt_U_0 / sqrt(3)
        self.data.update({"num_dt_U_0": num_dt_U_0, "num_u_U_0": num_u_U_0})
        # U_x不确定度
        num_dt_U_x = InstrumentError.dc_potentiometer(0.01, self.data['num_U_x'], 0.1)
        num_u_U_x = num_dt_U_x / sqrt(3)
        self.data.update({"num_dt_U_x": num_dt_U_x, "num_u_U_x": num_u_U_x})
        # 合成最终的不确定度
        num_u_Rx_Rx = sqrt((num_u_R_0 / self.data['num_R_0']) ** 2 + (num_u_U_x / self.data['num_U_x']) ** 2 + (num_u_U_0 / self.data['num_U_0']) ** 2)
        num_u_R_x = num_u_Rx_Rx * self.data['num_R_x']
        self.data.update({"num_u_Rx_Rx": num_u_Rx_Rx, "num_u_R_x": num_u_R_x})
        # 输出带不确定度的最终结果: 不确定度保留一位有效数字, 物理量结果与不确定度首位有效数字对齐
        self.data['final_2'] = Method.final_expression(self.data['num_R_x'], self.data['num_u_R_x'])
        print(self.data['final_2'])

    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 实验一
        # 表格：原始数据
        self.report_data['t'] = "%.2f" % self.data['num_t']
        for i, R_11_i in enumerate(self.data['list_R_11']):
            self.report_data[str(i + 1)] = "%.1f" % (R_11_i)
        for i, R_21_i in enumerate(self.data['list_R_21']):
            self.report_data[str(i + 4)] = "%.1f" % (R_21_i)
        for i, R_12_i in enumerate(self.data['list_R_12']):
            self.report_data[str(i + 7)] = "%.1f" % (R_12_i)
        for i, R_22_i in enumerate(self.data['list_R_22']):
            self.report_data[str(i + 10)] = "%.1f" % (R_22_i)
        for i, R_13_i in enumerate(self.data['list_R_13']):
            self.report_data[str(i + 13)] = "%.1f" % (R_13_i)
        for i, R_23_i in enumerate(self.data['list_R_23']):
            self.report_data[str(i + 16)] = "%.1f" % (R_23_i)
        # 表格：求平均值
        self.report_data['R_11_avg'] = "%.1f" % self.data['num_R_11_avg']
        self.report_data['R_21_avg'] = "%.1f" % self.data['num_R_21_avg']
        self.report_data['R_12_avg'] = "%.1f" % self.data['num_R_12_avg']
        self.report_data['R_22_avg'] = "%.1f" % self.data['num_R_22_avg']
        self.report_data['R_13_avg'] = "%.1f" % self.data['num_R_13_avg']
        self.report_data['R_23_avg'] = "%.1f" % self.data['num_R_23_avg']
        # 最终结果
        self.report_data['final_1'] = self.data['final_1']
        # 将各个变量以及不确定度的结果导入实验报告，在实际编写中需根据实验报告的具体要求设定保留几位小数
        self.report_data['E_x'] = "%.4f" % self.data['num_E_x']
        self.report_data['dt_R_11'] = "%.3f" % self.data['num_dt_R_11']
        self.report_data['dt_R_21'] = "%.3f" % self.data['num_dt_R_21']
        self.report_data['dt_R_12'] = "%.3f" % self.data['num_dt_R_12']
        self.report_data['dt_R_22'] = "%.3f" % self.data['num_dt_R_22']
        self.report_data['u_R_11'] = "%.3f" % self.data['num_u_R_11']
        self.report_data['u_R_21'] = "%.3f" % self.data['num_u_R_21']
        self.report_data['u_R_12'] = "%.3f" % self.data['num_u_R_12']
        self.report_data['u_R_22'] = "%.3f" % self.data['num_u_R_22']
        self.report_data['S'] = "%.1f" % self.data['num_S']
        self.report_data['dt_S'] = "%.8f" % self.data['num_dt_S']
        self.report_data['u_S'] = "%.6f" % self.data['num_u_S']
        self.report_data['u_Ex_Ex'] = "%.6f" % self.data['num_u_Ex_Ex']
        self.report_data['u_Ex'] = "%.6f" % self.data['num_u_Ex']
        # 实验二
        # 表格：原始数据
        self.report_data['R_0'] = "%d" % self.data['num_R_0']
        self.report_data['U_0'] = "%.6f" % self.data['num_U_0']
        self.report_data['U_x'] = "%.6f" % self.data['num_U_x']
        # 最终结果
        self.report_data['final_2'] = self.data['final_2']
        # 将各个变量以及不确定度的结果导入实验报告，在实际编写中需根据实验报告的具体要求设定保留几位小数
        self.report_data['R_x'] = "%.2f" % self.data['num_R_x']
        self.report_data['dt_R_0'] = "%.3f" % self.data['num_dt_R_0']
        self.report_data['u_R_0'] = "%.4f" % self.data['num_u_R_0']
        self.report_data['dt_U_0'] = "%.6f" % self.data['num_dt_U_0']
        self.report_data['u_U_0'] = "%.7f" % self.data['num_u_U_0']
        self.report_data['dt_U_x'] = "%.6f" % self.data['num_dt_U_x']
        self.report_data['du_U_x'] = "%.7f" % self.data['num_u_U_x']
        self.report_data['u_Rx_Rx'] = "%.6f" % self.data['num_u_Rx_Rx']
        self.report_data['u_R_x'] = "%.2f" % self.data['num_u_U_x']
        # 调用ReportWriter类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    ptm = Potentiometer()
