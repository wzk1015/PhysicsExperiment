import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
from numpy import sqrt, abs

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method
from reportwriter.ReportWriter import ReportWriter



class Michelson:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = [
        "1", "2", "3", "4", "5", "6", "7", "8", "9", "10",
        "5d-1", "5d-2", "5d-3", "5d-4", "5d-5", # 逐差法；5Δd
        "N", "d", "lbd", # N是100；每100个圈的光程差；光的波长
        "ua_d", "ub_d", "u_d",  # 100圈光程差d的不确定度
        "u_N", "u_lbd_lbd", "u_lbd",  # 不确定度的合成
        "final" # 最终结果
    ]

    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "data.xlsx"
    REPORT_TEMPLATE_FILENAME = "Spectrometer_empty.docx"
    REPORT_OUTPUT_FILENAME = "1071Report.docx"

    def __init__(self):
        self.data = {}  # 存放实验中的各个物理量
        self.uncertainty = {}  # 存放物理量的不确定度
        self.report_data = {}  # 存放需要填入实验报告的
        print("1091 迈克尔逊干涉仪实验\n1. 实验预习\n2. 数据处理")
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
            self.input_data("./"+self.DATA_SHEET_FILENAME)  # './' is necessary when running this file, but should be removed if run main.py
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
        ws = xlrd.open_workbook(filename).sheet_by_name('Michelson')
        # 从excel中读取数据
        list_d = []
        for row in [2, 4]:
            for col in range(1, 6):
                list_d.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['list_d'] = list_d # 存储从表格中读入的数据
    
    '''
    进行数据处理
    由于1091实验的数据处理非常非常简单，为节约代码量，将全部数据处理放在一个函数内完成.
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    def calc_data(self):
        num_N = 100
        self.data['num_N'] = num_N
        # 逐差法计算100圈光程差d
        list_dif_d, num_d = Method.successive_diff(self.data['list_d'])
        self.data['list_dif_d'] = list_dif_d
        self.data['num_d'] = num_d
        # 按公式计算待测光的波长
        num_lbd = 2 * num_d / (len(list_dif_d) * num_N)
        num_lbd = num_lbd * 1e6
        self.data['num_lbd'] = num_lbd
    '''
    计算所有的不确定度
    '''
    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    def calc_uncertainty(self):
        # 计算光程差d的a,b及总不确定度
        num_ua_d = Method.a_uncertainty(self.data['list_d']) # 这里容易写错，一定要用原始数据的数组
        num_ub_d = 0.00005 / sqrt(3)
        num_u_d = sqrt(num_ua_d ** 2 + num_ub_d ** 2)
        self.data.update({"num_ua_d":num_ua_d, "num_ub_d":num_ub_d, "num_u_d":num_u_d})
        # 计算圈数N的不确定度
        num_N = self.data['num_N']
        num_u_N = 1 / sqrt(3)
        self.data['num_u_N'] = num_u_N
        num_d, num_N = self.data['num_d'], self.data['num_N']
        # 波长的不确定度合成
        num_u_lbd_lbd = sqrt((num_u_d / num_d) ** 2 + (num_u_N / num_N) ** 2)
        num_lbd = self.data['num_lbd']
        num_u_lbd = num_u_lbd_lbd * num_lbd
        self.data.update({"num_u_lbd_lbd": num_u_lbd_lbd, "num_u_lbd": num_u_lbd})
        # 输出带不确定度的最终结果: 不确定度保留一位有效数字, 物理量结果与不确定度首位有效数字对齐
        # TODO: 输出最终结果的修约方法稍稍有点问题，待修改
        num_u_lbd_1bit, pwr = Method.scientific_notation(num_u_lbd) # 将不确定度转化为只有一位有效数字的科学计数法
        num_lbd_final = int(num_lbd * (10 ** pwr)) / (10 ** pwr) # 对物理量保留有效数字，截断处理
        self.data['final'] = "%.0f±%.0f" % (num_lbd_final, num_u_lbd_1bit) # 这里稍微有点问题，待修改......
    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 表格：原始数据d
        for i, d_i in enumerate(self.data['list_d']):
            self.report_data[str(i + 1)] = "%.5f" % (d_i) # 一定都是字符串类型
        # 表格：逐差法计算5Δd
        for i, dif_d_i in enumerate(self.data['list_dif_d']):
            self.report_data["5d-%d" % (i + 1)] = "%.5f" % (dif_d_i)
        # 最终结果
        self.report_data['final'] = self.data['final']
        # 将各个变量以及不确定度的结果导入实验报告，在实际编写中需根据实验报告的具体要求设定保留几位小数
        self.report_data['N'] = "%d" % self.data['num_N']
        self.report_data['d'] = "%.5f" % self.data['num_d']
        self.report_data['lbd'] = "%.2f" % self.data['num_lbd']
        self.report_data['ua_d'] = "%.5f" % self.data['num_ua_d']
        self.report_data['ub_d'] = "%.5f" % self.data['num_ub_d']
        self.report_data['u_d'] = "%.5f" % self.data['num_u_d']
        self.report_data['u_N'] = "%.5f" % self.data['num_u_N']
        self.report_data['u_lbd_lbd'] = "%.5f" % self.data['num_u_lbd_lbd']
        self.report_data['u_lbd'] = "%.5f" % self.data['num_u_lbd']
        # 调用ReportWriter类
        RW = ReportWriter()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    mc = Michelson()
