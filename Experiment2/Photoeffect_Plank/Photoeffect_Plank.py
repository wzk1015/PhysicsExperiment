# 光电效应测定普朗克常数

import xlrd
import os, sys, shutil
from numpy import array, asarray, abs
sys.path.append("../..")
from GeneralMethod.PyCalcLib import Method, Fitting
from GeneralMethod.Report import Report

class Photoeffect_Plank :
    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "Plank_data.xlsx"
    REPORT_TEMPLETE_FILENAME = "Plank_empty.docx"
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment2/2041Report.docx"

    def __init__(self):
        self.data = {}
        self.report_data = {}
        print("2041 光电效应测定普朗克常数")
        print("读入数据")
        self.input_data(self.DATA_SHEET_FILENAME) # OK
        # print(self.data)
        print("计算中......")
        self.calc_data()
        print("正在生成报告......")
        self.write_report()
        
        print("Done!")

    
    def input_data(self, filename):
        ws = xlrd.open_workbook(filename).sheet_by_index(0)
        # solve plank constant
        arr_lbd = []
        arr_nu = []
        arr_U0_auto = []
        arr_U0_man = []
        self.data['phi'] = float(ws.cell_value(2, 2))
        for col in range(2, 2 + 5):
            arr_lbd.append(ws.cell_value(5, col))
            arr_nu.append(ws.cell_value(6, col))
            arr_U0_man.append(ws.cell_value(7, col))
            arr_U0_auto.append(ws.cell_value(8, col))
        self.data['arr_lbd'] = arr_lbd
        self.data['arr_nu'] = arr_nu
        self.data['arr_U0_man'] = arr_U0_man
        self.data['arr_U0_auto'] = arr_U0_auto
        # U-I curve
        rbegin, rstep = 13, 4 # 第一个表格起始行 以及表格占的行数
        arr_UIcurve = []
        for i in range(4):
            rcur = rbegin + i * rstep # 当前表格的行数
            table_i = {} # 当前表格，包含一个波长和两列数据
            table_i['lambda'] = float(ws.cell_value(rcur, 2))
            table_i['arr_U'] = []
            table_i['arr_I'] = []
            for j in range(10):
                table_i['arr_U'].append(float(ws.cell_value(rcur + 2, j + 1)))
                table_i['arr_I'].append(float(ws.cell_value(rcur + 3, j + 1)))
            arr_UIcurve.append(table_i)
        self.data['arr_UIcurve'] = arr_UIcurve
        # finished

    def calc_data(self):
        arr_U0_auto = asarray(self.data['arr_U0_auto']) # 转换成numpy格式
        arr_U0_man = asarray(self.data['arr_U0_man'])
        arr_U0 = (arr_U0_auto + arr_U0_man) / 2 # TODO: 由于这里我不知道自动手动啥意思，就取平均了
        arr_nu = self.data['arr_nu']
        # 说明：这里跟北航学习生活圈的实验报告有些不同
        # 样例实验报告中的图示法是一种粗略的方法，只能手工作图
        # 如果用计算机的话，直接就是线性回归拟合了，因此只做第二种 线性回归法
        b, a, r = Fitting.linear(arr_nu, arr_U0) # 返回值顺序: 斜率, 截距, 相关系数
        b *= 1e-14
        print("b = %e, r = %g" % (b, r))
        self.data['a'] = a
        self.data['b'] = b
        self.data['r'] = r

        # TODO: 相对第一种方法的误差没法算，只能用第二种方法算相对误差
        std_e = 1.602176565e-19
        std_h = 6.62606957e-34
        self.data['h'] = b * std_e
        
        self.data['eta'] = abs(self.data['h'] - std_h) / std_h

        pass

    # 不用计算不确定度

    # TODO: 这个实验需要绘图，这个功能先暂时放一边（因为还没研究出来如何向word中插图）
    def draw_graph(self):
        
        pass


    def write_report(self):
        # 填常数
        self.report_data['phi'] = "%.0f" % self.data['phi']
        # 先填第一个表
        for i, lbd_i in enumerate(self.data['arr_lbd']):
            kw = "r-l%d" % (i + 1)
            self.report_data[kw] = "%.0f" % lbd_i
        for i, nu_i in enumerate(self.data['arr_nu']):
            kw = "r-v%d" % (i + 1)
            self.report_data[kw] = "%.3f" % nu_i
        for i, U0_man_i in enumerate(self.data['arr_U0_man']):
            kw = "r-um%d" % (i + 1)
            self.report_data[kw] = "%.3f" % U0_man_i
        for i, U0_auto_i in enumerate(self.data['arr_U0_auto']):
            kw = "r-ua%d" % (i + 1)
            self.report_data[kw] = "%.3f" % U0_auto_i
        self.report_data['s2-b'] = "%e" % (self.data['b'] * 1e-14)
        self.report_data['s2-r'] = "%.4f" % (self.data['r'])
        self.report_data['s2-h'] = "%e" % (self.data['h'])

        self.report_data['s12-eta'] = "%.1f" % (self.data['eta'] * 1e2)

        # 伏安特性曲线表格 填入
        for i in range(4):
            self.report_data['lbd%d' % (i + 1)] = self.data['arr_UIcurve'][i]['lambda']
            for j in range(10):
                kw = "%d-u%d" % (i + 1, j + 1)
                self.report_data[kw] = self.data['arr_UIcurve'][i]['arr_U'][j]
                kw = "%d-i%d" % (i + 1, j + 1)
                self.report_data[kw] = self.data['arr_UIcurve'][i]['arr_I'][j]
        
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLETE_FILENAME, self.REPORT_OUTPUT_FILENAME)
    pass

if __name__ == '__main__':
    Photoeffect_Plank()