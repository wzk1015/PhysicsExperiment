import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
import math
import sys
o_path = os.path.abspath(os.path.join(os.getcwd(), "../..")) # 调用库需要返回到当前目录的上上级
sys.path.append(o_path) # 如果最终要从main.py调用，则删掉这句

from GeneralMethod.PyCalcLib import Method, Fitting
from GeneralMethod.Report import Report


class Abel:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = [
        "1", "2", "3",  # 记录的间距
        "lambda", "F", # 波长和焦距
        "xi_1",	"xi_2",	"xi_3", # 计算得到的间距
        "fx_1",	"fx_2",	"fx_3",	# 计算得到的频率
        "f0", # 计算得到的基频
    ]
    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "data.xlsx"
    REPORT_TEMPLATE_FILENAME = "Abel_empty.docx"
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment2/2091Report.docx"

    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.report_data = {} # 存放需要填入实验报告的
        print("2091 阿贝尔效应和空间滤波实验\n1. 实验预习\n2. 数据处理")
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

    '''
    从excel表格中读取数据
    @param filename: 输入excel的文件名
    @return none
    '''
    def input_data(self, filename):
        # 从excel第一个工作簿中读取数据
        ws = xlrd.open_workbook(filename).sheet_by_name('Abel')       
        D = []
        for col in range(1, 4): 
            D.append(float(ws.cell_value(1, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        Lambda = float(ws.cell_value(5, 1))
        F = float(ws.cell_value(6, 1))
        self.data['D'] = D # 存储从表格中读入的数据
        self.data['lambda'] = Lambda
        self.data['F'] = F
    
    '''
    进行数据处理
    由于2091实验的数据处理非常非常简单，为节约代码量，将全部数据处理放在一个函数内完成.
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    def calc_data(self):
        # 计算1，2，3级衍射点与0级衍射点间距
        xi = []
        for i in range(0, 3):
            xi[i] = self.data['D'][i] * 5
        self.data['xi'] = xi
        #计算频率和基频
        fx = []
        f0 = 0.0
        for i in range(0, 3):
            fx[i] = xi[i] / (self.data['lambda'] * self.data['F']) * 1e9
            f0 += fx[i] / (i + 1)
        self.data['fx'] = fx
        self.data['f0'] = f0      

    '''
    本实验不需要计算不确定度
    '''

    '''
    填充实验报告
    调用Report类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 表格：原始数据
        for i, d_i in enumerate(self.data['D']):
            self.report_data[str(i + 1)] = "%.2f" % (d_i) # 一定都是字符串类型
        self.report_data['lambda'] = "%.1f" % self.data['lambda']
        self.report_data['F'] = "%d" % self.data['F']
        # 数据处理
        for i, xi_i in enumerate(self.data['xi']):
            self.report_data['xi_%d' % (i + 1)] = "%.2f" % (xi_i)
        for i, fx_i in enumerate(self.data['fx']):
            self.report_data['fx_%d' % (i + 1)] = "%.4f" % (fx_i)
        self.report_data['f0'] = "%.4f" % self.data['f0']
        # 调用Report类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    ab = Abel()
