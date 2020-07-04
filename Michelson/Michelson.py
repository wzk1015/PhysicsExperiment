import xlrd, xlwt
from xlutils.copy import copy as xlscopy
import shutil
import os
from numpy import sqrt, abs

import sys
sys.path.append('..')
#print(sys.path)
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

    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {}

        print("即将打开数据输入文件......")
        # os.startfile("data.xlsx")
        input("输入数据完成后请保存并关闭excel文件，然后按回车键继续")
        self.input_data("./data.xlsx") # './' is necessary when running this file, but should be removed if run main.py
        print("数据读入完毕，处理中......")
        self.calc_data()
        self.calc_uncertainty()
        # print(self.data)
        print("正在生成实验报告......")
        self.fill_report()
        print("Done!")

    '''
    从excel表格中读取数据
    @param filename: 输入excel的文件名
    @return none
    '''
    def input_data(self, filename):
        ws = xlrd.open_workbook(filename).sheet_by_name('Michelson')
        # 从excel中读取数据
        dlist = []
        for row in [2, 4]:
            for col in range(1, 6):
                dlist.append(float(ws.cell_value(row, col))) # 从excel取出来的数据，加个类型转换靠谱一点
        self.data['d_list'] = dlist # 存储从表格中读入的数据
    '''
    进行数据处理
    由于1091实验的数据处理非常非常简单，为节约代码量，将全部数据处理放在一个函数内完成.
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    def calc_data(self):
        N = 100
        self.data['N'] = N
        # 逐差法计算100圈光程差d
        dif_d, d = Method.successive_diff(self.data['d_list'])
        self.data['dif_d'] = dif_d
        self.data['d'] = d
        # 按公式计算待测光的波长
        lbd = 2 * d / (len(dif_d) * N)
        lbd = lbd * 1e6
        self.data['lbd'] = lbd
    '''
    计算所有的不确定度
    '''
    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    def calc_uncertainty(self):
        ua_d = Method.a_uncertainty(self.data['d_list']) # 这里容易写错，一定要用原始数据的数组
        ub_d = 0.00005 / sqrt(3)
        u_d = sqrt(ua_d ** 2 + ub_d ** 2)
        self.data.update({"ua_d":ua_d, "ub_d":ub_d, "u_d":u_d})
        N = self.data['N']
        u_N = 1 / sqrt(3)
        self.data['u_N'] = u_N
        d, N = self.data['d'], self.data['N']
        u_lbd_lbd = sqrt((u_d / d) ** 2 + (u_N / N) ** 2)
        lbd = self.data['lbd']
        u_lbd = u_lbd_lbd * lbd
        self.data.update({"u_lbd_lbd": u_lbd_lbd, "u_lbd": u_lbd})
        # 输出一下最终结果
        bse, pwr = Method.scientific_notation(u_lbd)
        lbd_f = int(lbd * (10 ** pwr)) / (10 ** pwr)
        self.data['final'] = "%.0f±%.0f" % (lbd_f, bse)
    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 原始数据d
        for i, di in enumerate(self.data['d_list']):
            # print(i + 1, di)
            self.report_data[str(i + 1)] = "%.5f" % (di) # 一定都是字符串类型
        # 逐差法计算5Δd
        for i, ddi in enumerate(self.data['dif_d']):
            self.report_data["5d-%d" % (i + 1)] = "%.5f" % (ddi)
        self.report_data['final'] = self.data['final']
        # self.report_data['lbd'] = "%.0f" % (self.data['lbd'] * 1e6)
        # 将各个变量以及不确定度的结果导入实验报告
        for key in self.report_data_keys:
            if not key in self.report_data.keys() and key in self.data.keys():
                self.report_data[key] = "%.5f" % (self.data[key])
        RW = ReportWriter()
        RW.load_replace_kw(self.report_data)
        RW.fill_report("Michelson_empty.docx", "Michelson_out.docx")

if __name__ == '__main__':
    mc = Michelson()
