import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
from numpy import sqrt, abs
import subprocess

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method,InstrumentError
from GeneralMethod.Report import Report


class Talbot:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含
    report_data_keys = [
        "itv"
        "1", "2", "3", "4", "5", "6"
        "z_m", "z_t", "eta"
    ]

    PREVIEW_FILENAME = "Preview.pdf"  # 预习报告模板文件的名称
    DATA_SHEET_FILENAME = "data.xlsx"  # 数据填写表格的名称
    REPORT_TEMPLATE_FILENAME = "Talbot_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment2/2191Report.docx"  # 最后生成实验报告的相对路径

    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {} # 存放需要填入实验报告的
        print("2191 双光栅实验\n1. 实验预习\n2. 数据处理")
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
        ws = xlrd.open_workbook(filename).sheet_by_name('Talbot')
        row = 1
        col = 1
        self.data['num_itv'] = int(ws.cell_value(row, col))
        list_d = []
        row = 2
        for col in range(1,7):
            list_d.append(float(ws.cell_value(row, col)))
        self.data['list_d'] = list_d
    '''
    进行数据处理

    '''
    def calc_data(self):
        list_d = self.data['list_d']
        num_itv = self.data['num_itv']
        num_z_m = (list_d[5] - list_d[0]) / (num_itv * 5)
        num_z_t = (0.01 ** 2) / (630 * 1e-6)
        self.data.update({"num_z_m": num_z_m, "num_z_t": num_z_t})

    '''
    计算所有的不确定度
    '''
    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    def calc_uncertainty(self):
        num_eta = abs((self.data['num_z_m'] - self.data['num_z_t']) / self.data['num_z_t']) 
        self.data['num_eta'] = num_eta

    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        for i, d_i in enumerate(self.data['list_d']):
            self.report_data[str(i + 1)] = "%.1f" % (d_i)
        self.report_data['itv'] = "%d" % self.data['num_itv']
        self.report_data['z_m'] = "%.3f" % self.data['num_z_m']
        self.report_data['z_t'] = "%.3f" % self.data['num_z_t']
        self.report_data['eta'] = "%.4f" % self.data['num_eta']
        # 调用ReportWriter类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    tbt = Talbot()
