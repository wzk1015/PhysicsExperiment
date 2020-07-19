import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
from numpy import sqrt, abs

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method
from reportwriter.ReportWriter import ReportWriter

class f_measure:
     # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = [
        #1061_1
        "u_1_1""u_1_2""u_1_3""u_1_4""v_1_1""v_1_2""v_1_3""v_1_4"
        "u_2_1""u_2_2""u_2_3""u_2_4""v_2_1""v_2_2""v_2_3""v_2_4"
        "f_1""f_2"
        #1061_2
        "x_1_1""x_1_2""x_1_3""x_1_4""x_1_5"
        "x_2_1""x_2_2""x_2_3""x_2_4""x_2_5"
        "list_f_1""list_f_2""list_f_3""list_f_4""list_f_5"
        "x""final_f""num_u_f"
        #1061_3
        "a_1""a_2""a_3""a_4""a_5"
        "b_1""b_2""b_3""b_4""b_5"
        "3_f_1""3_f_2""3_f_3""3_f_4""3_f_5"
        "overline_3_f""final_3_f""u_3_f"
        #1061_4




    ]

    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "data.xlsx"
    REPORT_TEMPLATE_FILENAME = "f_measure_empty.docx"
    REPORT_OUTPUT_FILENAME = "f_measure_out.docx"

    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.report_data = {} # 存放需要填入实验报告的
        print("1061  物距像距法测透镜焦距\n1. 实验预习\n2. 数据处理")
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
            self.calc_data_2()
            self.calc_data_3()
            #self.calc_data_4()待定
            # 计算不确定度
            self.calc_uncertainty_2()
            self.calc_uncertainty_3()
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            os.startfile(self.REPORT_OUTPUT_FILENAME)
            print("Done!")


    def input_data(self, filename):
        ws = xlrd.open_workbook(filename).sheet_by_name('f_measure')

        #1061_1
        list_u_1 = []
        list_u_2 = []
        list_v_1 = []
        list_v_2 = []

        for row in [3]:
            for col in range(1, 5):
                list_u_1.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_u_1'] = list_u_1 # 存储从表格中读入的数据
        for row in [8]:
            for col in range(1, 5):
                list_u_2.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_u_2'] = list_u_2 # 存储从表格中读入的数据
        for row in [4]:
            for col in range(1, 5):
                list_v_1.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_v_1'] = list_v_1 # 存储从表格中读入的数据
        for row in [9]:
            for col in range(1, 5):
                list_v_2.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_v_2'] = list_v_2 # 存储从表格中读入的数据

        #1061_2
        pl_x = 0
        list_x_1 = []
        list_x_2 = []

        pl_x = float(ws.cell_value(13, 1)) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['pl_x'] = pl_x # 存储从表格中读入的数据
        
        for row in [15]:
            for col in range(1, 6):
                list_x_1.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['x_1'] = list_x_1 # 存储从表格中读入的数据

        for row in [16]:
            for col in range(1, 6):
                list_x_2.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['x_2'] = list_x_2 # 存储从表格中读入的数据

        #1061_3
        list_a = []
        list_b = []

        for row in [21]:
            for col in range(1, 6):
                list_a.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_a'] = list_a # 存储从表格中读入的数据
        for row in [22]:
            for col in range(1, 6):
                list_b.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_b'] = list_b # 存储从表格中读入的数据

        #1061_4



        


    
   
    def calc_data_1(self):
         self.data['f1'] = f1
        list_f1 = [(u_1_1*v_1_1)/(u_1_1+v_1_1),(u_1_2*v_1_2)/(u_1_2+v_1_2),(u_1_3*v_1_3)/(u_1_3+v_1_3),(u_1_4*v_1_4)/(u_1_4+v_1_4)]
        f_1 = (f1[0]+f1[1]+f1[2]+f1[3])/4
         self.data['f2'] = f2
        list_f2 = [(u_2_1*v_2_1)/(u_2_1+v_2_1),(u_2_2*v_2_2)/(u_2_2+v_2_2),(u_2_3*v_2_3)/(u_2_3+v_2_3),(u_2_4*v_2_4)/(u_2_4+v_2_4)]
        f_2 = (f2[0]+f2[1]+f2[2]+f2[3])/4

    def calc_data_2(self):
        i = 0
        self.data['x_1'] = list_x_1 = []
        self.data['x_2'] = list_x_2 = []
        self.data["list_f"] = list_f = []
        while i<5 :
            list_f[i] = (list_x_1[i] + list_x_2[i])/2 - pl_x
            i = i + 1
        self.data["num_f"] = num_f = Method.average(self.data['list_f'])
        
    def calc_data_3(self):
        i = 0
        self.data['list_3_f'] = list_3_f = []
        self.data['list_a'] = list_a = []
        self.data['list_b'] = list_b = []
        self.data['overline_3_f'] = overline_3_f
        while i<5 :
            list_3_f[i] = ( list_a[i]^2 - list_b[i]^2 ) / 4*list_b[i]
            i = i + 1
        overline_3_f = Method.average(self.data['list_3_f'])
        

    # def calc_data_4(self): 待定



    def calc_uncertainty_2(self):
        self.data["num_u_f"] = num_u_f
        num_u_f = Method.a_uncertainty(self.data['list_f'])

    def calc_uncertainty_3(self):
        self.data["u_3_f"] = u_3_f
        u_3_f = Method.a_uncertainty(self.data['list_3_f'])

    
    def fill_report(self):
        
        #1061_1
        for i, list_u_1_i in enumerate(self.data['list_u_1']):
            self.report_data[u_1_(i + 1)] = "%.1f" % (list_u_1_i)
        for i, list_u_2_i in enumerate(self.data['list_u_2']):
            self.report_data[u_2_(i + 1)] = "%.1f" % (list_u_2_i)
        for i, list_v_1_i in enumerate(self.data['list_v_1']):
            self.report_data[v_1_(i + 1)] = "%.1f" % (list_v_1_i)
        for i, list_v_2_i in enumerate(self.data['list_v_2']):
            self.report_data[v_2_(i + 1)] = "%.1f" % (list_v_2_i)      
        self.report_data['f_1'] = "%.1f" % f_1
        self.report_data['f_2'] = "%.1f" % f_2
        for i, list_f1_i in enumerate(self.data['list_f1']):
            self.report_data[f1_(i + 1)] = "%.1f" % (list_f1_i)
        for i, list_f2_i in enumerate(self.data['list_f2']):
            self.report_data[f2_(i + 1)] = "%.1f" % (list_f2_i)

        #1061_2
        for i, list_x_1_i in enumerate(self.data['x_1']):
            self.report_data[x_1_(i + 1)] = "%.1f" % (list_x_1_i)
        for i, list_x_2_i in enumerate(self.data['x_2']):
            self.report_data[x_2_(i + 1)] = "%.1f" % (list_x_2_i)
        for i, list_f_i in enumerate(self.data['list_f']):
            self.report_data[list_f_(i + 1)] = "%.1f" % (list_f_i)
        self.report_data['num_u_f'] = "%.3f" % self.data['num_u_f']
        self.report_data["final_f"] =  "%.1f±%.1f" % (num_f, num_u_f)    

        #1061_3
        for i, list_a_i in enumerate(self.data['list_a']):
            self.report_data[a_(i + 1)] = "%.1f" % (list_a_i)
        for i, list_b_i in enumerate(self.data['list_b']):
            self.report_data[b_(i + 1)] = "%.1f" % (list_b_i)
        for i, list_3_f_i in enumerate(self.data['list_3_f']):
            self.report_data[3_f_(i + 1)] = "%.1f" % (list_3_f_i)
        self.report_data['final_3_f'] = "%.1f" % self.data['final_3_f']
        self.report_data['u_3_f'] = "%.3f" % self.data['u_3_f']

        #1061_4

        

        # 调用ReportWriter类
        RW = ReportWriter()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)



if __name__ == '__main__':
    fms = f_measure()