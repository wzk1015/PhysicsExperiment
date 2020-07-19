import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
from numpy import sqrt, abs

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method
from reportwriter.ReportWriter import ReportWriter



class Parallel:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = [
        #1062_1
            "f_0" "delta_yi"
            "y_1_1""y_1_2""y_1_3""y_1_4""y_1_5""y_1_6""y_1_7""y_1_8"
            "f_1_1""f_1_2""f_1_3""f_1_4""f_1_5""f_1_6""f_1_7""f_1_8"
            "u_1_1""u_1_2""u_1_3""u_1_4""u_1_5""u_1_6""u_1_7""u_1_8"
            "f_1" "u_1" "final_1"

        #1062_2 待定

    ]

    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "data.xlsx"
    REPORT_TEMPLATE_FILENAME = "Parallel_empty.docx"
    REPORT_OUTPUT_FILENAME = "Parallel_out.docx"

    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {} # 存放需要填入实验报告的
        print("1062 平行光管法测透镜焦距\n1. 实验预习\n2. 数据处理")
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
            self.calc_data_1()
            #self.calc_data_2() 待定

            # 计算不确定度
            self.calc_uncertainty_1()
            #self.calc_uncertainty_2() 待定 

            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            os.startfile(self.REPORT_OUTPUT_FILENAME)
            print("Done!")




    def input_data(self, filename):
        ws = xlrd.open_workbook(filename).sheet_by_name('Parallel')
        
        #1062_1
        f_0 = 0
        f_0 = (float(ws.cell_value(1, 1)))
        self.data['f_0'] = f_0

        list_y = []
        for row in [2]:
            for col in range(1, 9):
                list_y.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_y'] = list_y # 存储从表格中读入的数据

        list_y_1 = []
        for row in [3]:
            for col in range(1, 9):
                list_y_1.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_y_1'] = list_y_1 # 存储从表格中读入的数据

        delta_yi = 0
        delta_yi = (float(ws.cell_value(4, 1)))
        self.data['delta_yi'] = delta_yi

        #1062_2
    




    def calc_data_1(self):
        i = 0
        list_f_1 = []
        while i < 8:
            list_f_1[i] = (list_y_1 / list_y[i]) * f_0
            i = i+1
        self.data['list_f_1'] = list_f_1
        
        
        

    
    #def calc_data_2(self):待定




    def calc_uncertainty_1(self):
        u_b = 0
        u_b = self.data['delta_yi'] / sqrt(3)
        self.data['u_b'] = u_b
        
        i = 0
        list_u = []
        while i < 8:
            list_u[i] = (u_b / list_y[i]) * f_0
            i = i+1
        self.data['list_u'] = list_u

        i = 0
        f_1_a = 0
        for i < 8 :
            f_1_a = f_1_a + (list_f_1[i] / list_u[i] ^2)
            i = i+1

        i = 0
        f_1_b = 0
        for i < 8 :
            f_1_b = f_1_b + (1 / list_u[i] ^2)
            i = i+1
        
        f_1 = f_1_a / f_1_b
        u_1 = sqrt(1 / f_1_b)

        self.data['f_1'] = f_1
        self.data['u_1'] = u_1        
        self.data['final_1'] = "%.1f±%.1f" % (f_1, u_1)

        
    #def calc_uncertainty_2(self):待定




    def fill_report(self):
        
        #1062_1
        for i, list_y_1_i in enumerate(self.data['list_y_1']):
            self.report_data['y_1_(i + 1)'] = "%.3f" % (list_y_1_i) # 一定都是字符串类型
        for i, list_f_1_i in enumerate(self.data['list_f_1']):
            self.report_data['f_1_(i + 1)'] = "%.3f" % (list_f_1_i) # 一定都是字符串类型
        for i, list_u_i in enumerate(self.data['list_u']):
            self.report_data['u_1_(i + 1)'] = "%.3f" % (list_u_i) # 一定都是字符串类型
        
        self.report_data['final_1'] = "%.1f" % self.data['final_1']
        self.report_data['f_0'] = "%.3f" % self.data['f_0']
        self.report_data['delta_yi'] = "%.5f" % self.data['delta_yi']
        self.report_data['f_1'] = "%.3f" % self.data['f_1']
        self.report_data['u_1'] = "%.3f" % self.data['u_b']

        #1062_2

       
        # 调用ReportWriter类
        RW = ReportWriter()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    pms = Parallel()
