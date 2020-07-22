import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
from numpy import sqrt, abs, sin, asin

import sys
sys.path.append('..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method
from reportwriter.ReportWriter import ReportWriter


class Velocity:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = [
        "f_s_1""f_s_2""f_s_3""f_s_4""f_s_5""f_s_6""f_s_7""f_s_8"
        "phi_0_1""phi_0_2""phi_0_3""phi_0_4""phi_0_5""phi_0_6""phi_0_7""phi_0_8"
        "phi_1_1""phi_1_2""phi_1_3""phi_1_4""phi_1_5""phi_1_6""phi_1_7""phi_1_8"
        "b" "v_s" "percent"

    ]
    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "data.xlsx"
    REPORT_TEMPLATE_FILENAME = "Velocity_empty.docx"
    REPORT_OUTPUT_FILENAME = "Velocity_out.docx"

    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.report_data = {} # 存放需要填入实验报告的
        print("2201 声光衍射\n1. 实验预习\n2. 数据处理")
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
            os.startfile(self.PREVIEW_FILENAME)
        elif oper == '2':
            print("现在开始数据处理")
            print("即将打开数据输入文件......")
            # 打开数据输入文件
            os.startfile(self.DATA_SHEET_FILENAME)
            input("输入数据完成后请保存并关闭excel文件，然后按回车键继续")
            # 从excel中读取数据
            self.input_data("./"+self.DATA_SHEET_FILENAME) # './' is necessary when running this file, but should be removed if run main.py
            print("数据读入完毕，处理中......")
            # 计算物理量
            self.calc_data()
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            os.startfile(self.REPORT_OUTPUT_FILENAME)
            print("Done!")

    


    def input_data(self, filename):
        ws = xlrd.open_workbook(filename).sheet_by_name('Velocity')
        list_f_s = []
        for row in [3]:
            for col in range(2, 10):
                list_f_s.append(float(ws.cell_value(row, col))) 
        self.data['list_f_s'] = list_f_s 
    
        list_phi_0 = []
        for row in [4]:
            for col in range(2, 10):
                list_phi_0.append(float(ws.cell_value(row, col))) 
        self.data['list_phi_0'] = list_phi_0
    
        num_n = 0
        num_n = float(ws.cell_value(5, 2))
        self.data['num_n'] = num_n

        num_lambda_0 = 0
        num_lambda_0 = float(ws.cell_value(6, 2))
        self.data['num_lambda_0'] = num_lambda_0

        num_v_0 = 0
        num_v_0 = float(ws.cell_value(7, 2))
        self.data['num_v_0'] = num_v_0
    

    
    def calc_data(self):
        list_phi_1 = []
        i = 0
        while i < 8:
            list_phi_1[i] = asin(sin(list_phi_0[i]) / num_n)
            i = i + 1
        self.data['list_phi_1'] = list_phi_1

        list_num_b = []
        list_num_b = Method.linear(list_f_s , list_phi_1)
        num_b = list_num_b[0]
        self.data['num_b'] = num_b

        num_v_s = num_lambda_0 * num_b / num_n 
        self.data[' num_v_s'] =  num_v_s
        num_percent = abs( num_v_s - num_v_0 ) / num_v_0 * 100
        self.data[' num_percent'] =  num_percent

    
    def fill_report(self):
        
        for i, list_f_s_i in enumerate(self.data['list_f_s']):
            self.report_data['f_s_(i + 1)'] = "%.2f" % (list_f_s_i)
        for i, list_phi_0_i in enumerate(self.data['list_phi_0']):
            self.report_data['phi_0_(i + 1)'] = "%.4f" % (list_phi_0_i)
        for i, list_phi_1_i in enumerate(self.data['list_phi_1']):
            self.report_data['phi_1_(i + 1)'] = "%.2f" % (list_phi_1_i)
        
        self.report_data['b'] =  self.data['num_b']
        self.report_data['v_s'] ="%.2f" % self.data['num_v_s']
        self.report_data['percent'] = "%.2f" % self.data['num_percent']
        


        RW = ReportWriter()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    vms = Velocity()
