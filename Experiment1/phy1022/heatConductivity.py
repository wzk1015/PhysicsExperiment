import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
import math
from numpy import sqrt, abs
import pandas
import sys
import numpy
import matplotlib
import scipy
# sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
sys.path.append('./.')
from GeneralMethod.PyCalcLib import Method
from GeneralMethod.Report import Report

class heatConductivity:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = [
        "list_q0", "list_q1",
        "pl_t",
        "list_v_10", "list_v_11", "list_v_12", "list_v_13", "list_v_14", "list_v_15", "list_v_16", "list_v_17", "list_v_18", "list_v_19",
        "pl_m", 
        "db", "dp", "hb", "hp",
        "d_ba", "d_bb", "d_bc",
        "h_ba", "h_bb", "h_bc",
        "d_pa", "d_pb", "d_pc",
        "h_pa", "h_pb", "h_pc", 
        "db", "dp", "hp", "hb", #平均值#
        "v", # 散热速率
        "k", # 冷却率
        "u_m", "ua_hp","ua_db", "ua_dp", "ua_hb", "ub_hp", "ub_hb", "ub_db", "hb_hp", # 不确定度
        "u_db", "u_hb", "u_dp", "u_hp", "u_kk", "uk", # 最终的不确定度
        "final" # 最终结果
    ]

    PREVIEW_FILENAME = "E:\基物实验程序\PhysicsExperiment\Experiment1\phy1022\Preview.pdf"
    #"Preview.pdf"
    DATA_SHEET_FILENAME = "E:\基物实验程序\PhysicsExperiment\Experiment1\phy1022\data.xlsx"
    #"E:\基物实验程序\PhysicsExperiment\Experiment1\phy1022\heatConductivity.xlsx"
    REPORT_TEMPLATE_FILENAME = "E:\基物实验程序\PhysicsExperiment\Experiment1\phy1022\heatConductivity_empty.docx"
    REPORT_OUTPUT_FILENAME = "E:\基物实验程序\PhysicsExperiment\Experiment1\phy1022\heatConductivity_out.docx"

    print(os.path.abspath(DATA_SHEET_FILENAME))
    Method.start_file(DATA_SHEET_FILENAME)
    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {} # 存放需要填入实验报告的
        print("1022 稳态法测不良导体热导率\n1. 实验预习\n2. 数据处理")
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
            self.input_data(self.DATA_SHEET_FILENAME) # './' is necessary when running this file, but should be removed if run main.py
            print("数据读入完毕，处理中......")
            # 计算物理量
            self.calc_data()
            # 计算不确定度
            self.calc_uncertainty()
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            os.startfile(self.REPORT_OUTPUT_FILENAME)
            print("Done!")


    '''
    从excel表格中读取数据
    @param filename: 输入excel的文件名
    @return none
    '''
    def input_data(self, filename):
       # print("initialyes")
        ws = xlrd.open_workbook(filename).sheet_by_name('Sheet1')
       # print("xlsx after yes")
        # 从excel中读取数据
        list_q = []
        pl_m = 0
        list_t = []
        list_v_1 = []
        list_db = []
        list_hb = []
        list_dp = []
        list_hp = []
       # print("YES")
        for row in [2, 3]:
            for col in [2]:
                print("%s" % ws.cell_value(row, col))
                list_q.append(ws.cell_value(row, col))
        self.data['list_q'] = list_q
        for row in [6]:
            for col in [2, 12]:
                list_q.append(float(ws.cell_value(row, col)))
        self.data['list_t'] = list_t

        for row in [7]:
            for col in [2, 12]:
                list_v_1.append(float(ws.cell_value(row, col)))
        self.data['list_v_1'] = list_v_1
        
        pl_m = float(ws.cell_value(10, 2))
        self.data['pl_m'] = pl_m
        
        for row in [11]:
            for col in [2,4]:
                list_db.append(float(ws.cell_value(row, col)))
        self.data['list_db'] = list_db # 存储从表格中读入的数据
        for row in [12]:
            for col in [2,4]:
                list_hb.append(float(ws.cell_value(row, col)))
        self.data['list_hb'] = list_hb
        for row in [13]:
            for col in [2,4]:
                list_dp.append(float(ws.cell_value(row, col)))
        self.data['list_dp'] = list_dp
        for row in [14]:
            for col in [2,4]:
                list_hp.append(float(ws.cell_value(row, col)))
        self.data['list_hp'] = list_hp
    '''
    进行数据处理
    '''

    def calc_data(self):
        x = numpy.array([self.data['list_t']])
        y = numpy.array([self.data['list_v_1']])
        xx = numpy.linspace(0, 10*self.data['list_t[1]'], 100)
        matplotlib.pyplot.scatter(x, y)
        f = scipy.interpolate.interp1d(x, y, kind = "cubic")
        ynew = f(xx)
        matplotlib.pyplot.plot(xx, ynew, "g")
        matplotlib.pyplot.show()
        ynew1 = f(self.data['list_q[1]'] + 0.001)
        ynew2 = f(self.data['list_q[1]'] - 0.001)
        v = math.fabs(ynew1 - ynew2) / 0.002
        self.data['v'] = v

        db = Method.average(self.data['list_db'])
        self.data['db'] = db
        dp = Method.average(self.data['list_dp'])
        self.data['dp'] = dp
        hb = Method.average(self.data['list_hb'])
        self.data['hb'] = hb
        hp = Method.average(self.data['list_hp'])
        self.data['hp'] = hp
        k = self.data['pl_m'] * 3 * 10**8 * self.data['v'] * (4*self.data['hp'] + self.data['dp']) / (self.data['dp'] + 2*self.data['hp']) * self.data['hb'] / (self.data['list_q[0]'] -self.data['list_q[1]']) * 2 / (math.pi * self.data['db'] ** 2)
        self.data['k'] = k

    def calc_uncertainty(self):
        u_m = self.data['pl_m'] / math.sqrt(3)
        self.data['u_m'] = u_m
        ua_hp = self.calc_uncertainty_1(self.data['list_hp'], self.data['hp'])
        self.data['ua_hp'] = ua_hp
        ua_db = self.calc_uncertainty_1(self.data['list_db'], self.data['db'])
        self.data['ua_db'] = ua_db
        ua_dp = self.calc_uncertainty_1(self.data['list_dp'], self.data['dp'])
        self.data['ua_dp'] = ua_dp
        ua_hb = self.calc_uncertainty_1(self.data['list_hb'], self.data['hb'])
        self.data['ua_hb'] = ua_hb
        ub_hp = ub_db = ub_dp = ub_hb = 0.02/math.sqrt(3)
        self.data['ub_hp'] = self.data['ub_db'] = self.data['ub_dp'] = self.data['ub_hb'] = ua_hp
        u_db = math.sqrt(ua_db**2 + ub_db**2)
        self.data['u_db'] = u_db
        u_hp = math.sqrt(ua_hp**2 + ub_hp**2)
        self.data['u_hp'] = u_hp
        u_dp = math.sqrt(ua_dp**2 + ub_dp**2)
        self.data['u_dp'] = u_dp
        u_hb = math.sqrt(ua_hb**2 + ub_hb**2)
        self.data['u_hb'] = u_hb
        u_kk = math.sqrt((self.data['u_m']/self.data['pl_m'])**2 + ((1/(self.data['dp']+4*self.data['hp'])-1/(self.data['dp']+2*self.data['hp'])*self.data['u_dp'])**2) + ((4/(self.data['dp']+4*self.data['hp'])-2/(self.data['dp']+2*self.data['hp']))*self.data['u_hp'])**2 + (self.data['u_hb']/self.data['hb'])**2 + (2*self.data['u_db']/self.data['db'])**2)
        self.data['u_kk'] = u_kk
        self.data['k'] = u_kk * self.data['k']

    def calc_uncertainty_1(self, list, a):
        temp = math.sqrt(((list[0] - a)**2 + (list[1] - a)**2 + (list[2] - a)**2)/6)
        return temp
    
    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 表格：原始数据d
        print('YES')
        self.report_data[str(1)] = "%.5f" % (self.data['v'])
        self.report_data[str(2)] = "%.5f" % (self.data['k'])
        self.report_data[str(3)] = "%.5f" % (self.data['u_m'])
        self.report_data[str(4)] = "%.5f" % (self.data['ua_db'])
        self.report_data[str(5)] = "%.5f" % (self.data['ua_dp'])
        self.report_data[str(6)] = "%.5f" % (self.data['ua_hp'])
        self.report_data[str(7)] = "%.5f" % (self.data['ua_hb'])
        self.report_data[str(8)] = "%.5f" % (self.data['ub_hp'])
        self.report_data[str(9)] = "%.5f" % (self.data['ub_db'])
        self.report_data[str(10)] = "%.5f" % (self.data['ub_dp'])
        self.report_data[str(11)] = "%.5f" % (self.data['ub_hb'])

        self.report_data['v'] = self.data['v']
        self.report_data['k'] = self.data['k']
        self.report_data['u_m'] = self.data['u_m']
        self.report_data['ua_hp'] = self.data['ua_hp']
        self.report_data['ua_db'] = self.data['ua_db']
        self.report_data['ua_hb'] = self.data['ua_hb']
        self.report_data['ua_dp'] = self.data['ua_dp']
        self.report_data['ub_dp'] = self.data['ub_dp']
        self.report_data['ub_hp'] = self.data['ub_hp']
        self.report_data['ub_bp'] = self.data['ub_bp']
        self.report_data['ub_db'] = self.data['ub_db']
        self.report_data['uk'] = self.data['uk']
        self.report_data['final'] = self.data['final']
        print("report")
        # 调用Report类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    hc = heatConductivity()