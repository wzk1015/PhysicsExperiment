import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import math
import os
from numpy import sqrt, abs

import sys
sys.path.append("..") # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import InstrumentError, Method, Fitting
from GeneralMethod.PyCalcLib import Report

class Kelvinbridge:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = [
        # 双电桥测电阻
        "R_1","R_N" #电阻R_1=R_2
        "1", "2", "3", "4", "5", "6", "7", "8", # 待测电阻长度
        "R-1", "R-2", "R-3", "R-4", "R-5", "R-6", "R-7", "R-8", # R_正
        "R-9", "R-10", "R-11", "R-12", "R-13", "R-14", "R-15", "R-16", # R_负
        "R_avg-1", "R_avg-2", "R_avg-3", "R_avg-4", "R_avg-5", "R_avg-6", "R_avg-7", "R_avg-8", # R_avg
        "R_x-1", "R_x-2", "R_x-3", "R_x-4", "R_x-5", "R_x-6", "R_x-7", "R_x-8", # R_x
        "D-1", "D-2", "D-3", "D-4", "D-5", "D-6", "D-7", "D-8", # 直径D
        "D", "b", "r", "rho" # 直径D的平均值，一元线性回归求得电阻率rho
        "u_b", "ua_D", "ub_D", "u_D"  # b和D的不确定度
        "u_rho_rho", "u_rho",  # 不确定度的合成
        "final_1" # 最终结果
        # 单电桥测电阻
        "R_3", "delta_n", "delta_R_3", "S", "a_pct", "R_0" #a_pct为误差系数,R_0为有效量程的基准值（该量程中最大的10的整数幂），S为灵敏度
        "delta_lmd", "u_lmd" # 灵敏度的不确定度
        "delta_yi", "u_yi" # 仪器的不确定度
        "u_R_x" # 不确定度的合成
        "final_2" # 最终结果
    ]

    PREVIEW_FILENAME = "Preview.pdf"  # 预习报告模板文件的名称
    DATA_SHEET_FILENAME = "data.xlsx"   # 数据填写表格的名称
    REPORT_TEMPLATE_FILENAME = "Kelvinbridge_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment1/1042Report.docx"  # 最后生成实验报告的相对路径


    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {} # 存放需要填入实验报告的
        print("1042 双电桥法测电阻实验\n1. 实验预习\n2. 数据处理")
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
        ws = xlrd.open_workbook(filename).sheet_by_name('1042')
        # 从excel中读取数据
        list_l = []
        row = 0
        for col in range(1, 9):
            list_l.append(int(ws.cell_value(row, col)) / 100) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_l'] = list_l # 存储从表格中读入的数据
        list_R = []
        for row in range(1,3):
            for col in range(1,9):
                list_R.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_R'] = list_R
        list_D = []
        row = 3
        for col in range(1, 9):
            list_D.append(float(ws.cell_value(row, col)) / 1000) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_D'] = list_D
        row = 4
        col = 1
        num_R_1 = int(ws.cell_value(row, col))
        self.data['num_R_1'] = num_R_1
        row = 5
        num_R_N = float(ws.cell_value(row, col))
        self.data['num_R_N'] = num_R_N
        row = 6
        num_R_3 = float(ws.cell_value(row, col))
        self.data['num_R_3'] = num_R_3
        row = 7
        num_delta_n = float(ws.cell_value(row, col))
        self.data['num_delta_n'] = num_delta_n
        row = 8
        num_delta_R_3 = float(ws.cell_value(row, col))
        self.data['num_delta_R_3'] = num_delta_R_3
        row = 9
        num_a_pct = float(ws.cell_value(row, col))
        self.data['num_a_pct'] = num_a_pct
        row = 10
        num_R_0 = float(ws.cell_value(row, col))
        self.data['num_R_0'] = num_R_0
        list_R_avg = []
        list_R_x = []
        for col in range(0,8):
            list_R_avg.append(0.5 * (list_R[col] + list_R[col+8]))
            list_R_x.append(num_R_N / num_R_1 * list_R_avg[col])
        self.data['list_R_avg'] = list_R_avg
        self.data['list_R_x'] = list_R_x

    '''
    进行数据处理
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    def calc_data(self):
        # 求直径D的平均值
        num_D = Method.average(self.data['list_D'])
        self.data['num_D'] = num_D
        # 一元线性回归求电阻率
        num_b, num_a, num_r = Fitting.linear(self.data['list_l'], self.data['list_R_x'])
        self.data['num_b'] = num_b
        self.data['num_r'] = num_r
        num_rho = math.pi * num_D * num_D * num_b / 4
        self.data['num_rho']=num_rho

    '''
    计算所有的不确定度
    '''
    def calc_uncertainty(self):
        # 计算一元线性回归的不确定度
        num_u_b = self.data['num_b'] * sqrt((1 / (self.data['num_r'] ** 2) - 1) / (8 - 2))
        self.data['num_u_b'] = num_u_b
        # 计算直径D的不确定度
        num_ua_D = Method.a_uncertainty(self.data['list_D'])
        num_ub_D = 0.03 / sqrt(3) /1000
        num_u_D = sqrt(num_ua_D ** 2 + num_ub_D ** 2)
        self.data.update({"num_ua_D":num_ua_D, "num_ub_D":num_ub_D, "num_u_D":num_u_D})
        # 电阻率不确定度的合成
        num_u_rho_rho = sqrt((num_u_D / self.data['num_D']) ** 2 + (num_u_b / self.data['num_b']) ** 2)
        num_rho = self.data['num_rho']
        num_u_rho = num_u_rho_rho * num_rho
        self.data.update({"num_u_rho_rho": num_u_rho_rho, "num_u_rho": num_u_rho})
        # 输出带不确定度的最终结果: 不确定度保留一位有效数字, 物理量结果与不确定度首位有效数字对齐
        self.data['final_1'] = Method.final_expression(self.data['num_rho'], self.data['num_u_rho'])
        # 单电桥测电阻
        # 灵敏度分析
        num_S = self.data['num_delta_n'] / self.data['num_delta_R_3']
        num_delta_lmd = 0.2 / num_S
        num_u_lmd = num_delta_lmd / sqrt(3)
        self.data.update({"num_S": num_S, "num_delta_lmd":num_delta_lmd, "num_u_lmd":num_u_lmd})
        # 仪器误差
        num_delta_yi = InstrumentError.dc_bridge(self.data['num_a_pct'], self.data['num_R_3'],self.data['num_R_0'])
        num_u_yi = num_delta_yi / sqrt(3)
        num_u_R_x = sqrt(num_u_yi ** 2 + num_u_lmd ** 2)
        self.data.update({"num_delta_yi": num_delta_yi, "num_u_yi": num_u_yi, "num_u_R_x": num_u_R_x})
        # 输出带不确定度的最终结果: 不确定度保留一位有效数字, 物理量结果与不确定度首位有效数字对齐
        self.data['final_2'] = Method.final_expression(self.data['num_R_3'], self.data['num_u_R_x'])
        print(self.data['final_2'])

    def fill_report(self):
        # 表格：原始数据
        self.report_data['R_1'] = "%d" % self.data['num_R_1']
        self.report_data['R_N'] = "%.5f" % self.data['num_R_N']
        self.report_data['R_3'] = "%.2f" % self.data['num_R_3']
        self.report_data['delta_n'] = "%d" % self.data['num_delta_n']
        self.report_data['delta_R_3'] = "%.2f" % self.data['num_delta_R_3']
        self.report_data['a_pct'] = "%.2f" % self.data['num_a_pct']
        self.report_data['R_0'] = "%.f" % self.data['num_R_0']
        for i, l_i in enumerate(self.data['list_l']):
            self.report_data[str(i + 1)] = "%d" % (l_i) # 一定都是字符串类型
        for i, R_i in enumerate(self.data['list_R']):
            self.report_data["R-%d" % (i + 1)] = "%.2f" % (R_i) 
        for i, R_avg_i in enumerate(self.data['list_R_avg']):
            self.report_data["R_avg-%d" % (i + 1)] = "%.3f" % (R_avg_i)
        for i, R_x_i in enumerate(self.data['list_R_x']):
            self.report_data["R_x-%d" % (i + 1)] = "%.7f" % (R_x_i)
        for i, D_i in enumerate(self.data['list_D']):
            self.report_data["D-%d" % (i + 1)] = "%.2f" % ((D_i) * 1000)
        # 最终结果
        self.report_data['final_1'] = self.data['final_1']
        self.report_data['final_2'] = self.data['final_2']
        # 将各个变量以及不确定度的结果导入实验报告，在实际编写中需根据实验报告的具体要求设定保留几位小数
        self.report_data['D'] = "%.3f" % (self.data['num_D'] * 1000)
        self.report_data['b'] = "%.4f" % self.data['num_b']
        self.report_data['r'] = "%.4f" % self.data['num_r']
        self.report_data['rho'] = "%.12f" % self.data['num_rho']
        self.report_data['u_b'] = "%.9f" % self.data['num_u_b']
        self.report_data['ua_D'] = "%.5f" % (self.data['num_ua_D'] * 1000)
        self.report_data['ub_D'] = "%.5f" % (self.data['num_ub_D'] * 1000)
        self.report_data['u_D'] = "%.5f" % self.data['num_u_D']
        self.report_data['u_rho_rho'] = "%.5f" % self.data['num_u_rho_rho']
        self.report_data['u_rho'] = "%.12f" % self.data['num_u_rho']
        self.report_data['S'] = "%.4f" % self.data['num_S']
        self.report_data['delta_lmd'] = "%.5f" % self.data['num_delta_lmd']
        self.report_data['u_lmd'] = "%.7f" % self.data['num_u_lmd']
        self.report_data['delta_yi'] = "%.5f" % self.data['num_delta_yi']
        self.report_data['u_yi'] = "%.5f" % self.data['num_u_yi']
        self.report_data['u_R_x'] = "%.5f" % self.data['num_u_R_x']
        # 调用ReportWriter类
        RW = ReportWriter()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)



if __name__ == '__main__':
    kb = Kelvinbridge()

