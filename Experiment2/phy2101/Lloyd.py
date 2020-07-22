import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
from numpy import sqrt, abs
import subprocess

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method
from GeneralMethod.Report import Report


class Lloyd:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = [
        "yl_1",	"yr_1",	"yd_1",
        "yl_2",	"yr_2",	"yd_2",
        "yl_3",	"yr_3",	"yd_3",
        "yl_4",	"yr_4",	"yd_4",
        "yl_5",	"yr_5",	"yd_5",
        "yl_6",	"yr_6",	"yd_6", "yd_avg", # 黄光的左右距离和间距
        "gl_1",	"gr_1",	"gd_1",
        "gl_2",	"gr_2",	"gd_2",
        "gl_3",	"gr_3",	"gd_3",
        "gl_4",	"gr_4",	"gd_4",
        "gl_5",	"gr_5",	"gd_5",
        "gl_6",	"gr_6",	"gd_6", "gd_avg", # 绿光的左右距离和间距
        "pl_1",	"pr_1",	"pd_1",
        "pl_2",	"pr_2",	"pd_2",
        "pl_3",	"pr_3",	"pd_3",
        "pl_4",	"pr_4",	"pd_4",
        "pl_5",	"pr_5",	"pd_5",
        "pl_6",	"pr_6",	"pd_6", "pd_avg", # 紫光的左右距离和间距
        "ll_1",	"lr_1",	"ld_1",
        "ll_2",	"lr_2",	"ld_2",
        "ll_3",	"lr_3",	"ld_3",
        "ll_4",	"lr_4",	"ld_4",
        "ll_5",	"lr_5",	"ld_5",
        "ll_6",	"lr_6",	"ld_6", "ld_avg", # 激光的左右距离和间距
        "lambda_y", "lambda_g", "lambda_p" # 波长
        "ua_d0", "u_d0", "ua_dy", "u_dy", "u_lambda_y","ua_dg", "u_dg", "u_lambda_g", "ua_dp", "u_dp", "u_lambda_p", # 不确定度
    ]

    PREVIEW_FILENAME = "Preview.pdf"  # 预习报告模板文件的名称
    DATA_SHEET_FILENAME = "data.xlsx"  # 数据填写表格的名称
    REPORT_TEMPLATE_FILENAME = "Lloyd_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment2/2101Report.docx"  # 最后生成实验报告的相对路径

    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {} # 存放需要填入实验报告的
        print("2101 劳埃镜的白光干涉实验\n1. 实验预习\n2. 数据处理")
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
        ws = xlrd.open_workbook(filename).sheet_by_name('Lloyd')
        # 从excel中读取数据
        yl = [], yr = [], gl = [], gr = [], pl = [], pr = [], ll = [], lr = []
        for row in range(2, 8):
            yl.append(float(ws.cell_value(row, 0)))  # 从excel取出来的数据，加个类型转换靠谱一点
            yr.append(float(ws.cell_value(row, 1)))
            gl.append(float(ws.cell_value(row, 2)))
            gr.append(float(ws.cell_value(row, 3)))
            pl.append(float(ws.cell_value(row, 4)))
            pr.append(float(ws.cell_value(row, 5)))
            ll.append(float(ws.cell_value(row, 6)))
            lr.append(float(ws.cell_value(row, 7)))
        lambda_0 = float(ws.cell_value(9, 1))
        self.data['yl'] = yl  # 存储从表格中读入的数据
        self.data['yr'] = yr
        self.data['gl'] = gl
        self.data['gr'] = gr
        self.data['pl'] = pl
        self.data['pr'] = pr
        self.data['ll'] = ll
        self.data['lr'] = lr
        self.data['lambda_0'] = lambda_0
    '''
    进行数据处理
    由于2101实验的数据处理非常非常简单，为节约代码量，将全部数据处理放在一个函数内完成.
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    def calc_data(self):
        # 计算间距
        yd = [], gd = [], pd = [], ld = []
        for i in range(0, 6):
            yd[i] = self.data['yl'][i] - self.data['yr'][i]
            gd[i] = self.data['gl'][i] - self.data['gr'][i]
            pd[i] = self.data['pl'][i] - self.data['pr'][i]
            ld[i] = self.data['ll'][i] - self.data['lr'][i]
        self.data['yd'] = yd, self.data['gd'] = gd, self.data['pd'] = pd, self.data['ld'] = ld
        yd_avg = Method.average(yd)
        gd_avg = Method.average(gd)
        pd_avg = Method.average(pd)
        ld_avg = Method.average(ld)
        self.data['yd_avg'] = yd_avg, self.data['gd_avg'] = gd_avg, self.data['pd_avg'] = pd_avg, self.data['ld_avg'] = ld_avg
        # 计算波长
        lambda_y = yd_avg / ld_avg * self.data['lambda_0']
        lambda_g = gd_avg / ld_avg * self.data['lambda_0']
        lambda_p = pd_avg / ld_avg * self.data['lambda_0']
        self.data['lambda_y'] = lambda_y, self.data['lambda_g'] = lambda_g, self.data['lambda_p'] = lambda_p

    '''
    计算所有的不确定度
    '''
    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    def calc_uncertainty(self):
        # 计算不确定度
        ub_d = 0.001 / sqrt(3)
        ua_d0 = Method.a_uncertainty(self.data['ld']) # 这里容易写错，一定要用原始数据的数组
        u_d0 = sqrt(ua_d0**2 + ub_d**2)
        ua_dy = Method.a_uncertainty(self.data['yd'])
        u_dy = sqrt(ua_dy**2 + ub_d**2)
        u_lambda_y = self.data['lambda_y'] * sqrt((u_dy / self.data['yd_avg'])**2 + (u_d0 / self.data['ld_avg'])**2)
        self.data.update({"ub_d":ub_d, "ua_d0":ua_d0, "u_d0":u_d0, "ua_dy":ua_dy, "u_dy":u_dy, "u_lambda_y":u_lambda_y})
        ua_dg = Method.a_uncertainty(self.data['gd'])
        u_dg = sqrt(ua_dg**2 + ub_d**2)
        u_lambda_g = self.data['lambda_g'] * sqrt((u_dg / self.data['gd_avg'])**2 + (u_d0 / self.data['ld_avg'])**2)
        ua_dp = Method.a_uncertainty(self.data['pd'])
        u_dp = sqrt(ua_dp**2 + ub_d**2)
        u_lambda_p = self.data['lambda_p'] * sqrt((u_dp / self.data['pd_avg'])**2 + (u_d0 / self.data['ld_avg'])**2)
        self.data.update({"ua_dg":ua_dg, "u_dg":u_dg, "u_lambda_g":u_lambda_g, "ua_dp":ua_dp, "u_dp":u_dp, "u_lambda_p":u_lambda_p})        
    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self): 
        # 表格：黄光
        for i, yl_i in enumerate(self.data['yl']):
            self.report_data["yl_%d" % (i + 1)] = "%.3f" % (yl_i) # 一定都是字符串类型
        for i, yr_i in enumerate(self.data['yr']):
            self.report_data["yr_%d" % (i + 1)] = "%.3f" % (yr_i)
        for i, yd_i in enumerate(self.data['yd']):
            self.report_data["yd_%d" % (i + 1)] = "%.3f" % (yd_i)
        self.report_data['yd_avg'] = self.data['yd_avg']
        # 表格：绿光
        for i, gl_i in enumerate(self.data['gl']):
            self.report_data["gl_%d" % (i + 1)] = "%.3f" % (gl_i)
        for i, gr_i in enumerate(self.data['gr']):
            self.report_data["gr_%d" % (i + 1)] = "%.3f" % (gr_i)
        for i, gd_i in enumerate(self.data['gd']):
            self.report_data["gd_%d" % (i + 1)] = "%.3f" % (gd_i)
        self.report_data['gd_avg'] = self.data['gd_avg']
        # 表格：紫光
        for i, pl_i in enumerate(self.data['pl']):
            self.report_data["pl_%d" % (i + 1)] = "%.3f" % (pl_i)
        for i, pr_i in enumerate(self.data['pr']):
            self.report_data["pr_%d" % (i + 1)] = "%.3f" % (pr_i)
        for i, pd_i in enumerate(self.data['pd']):
            self.report_data["pd_%d" % (i + 1)] = "%.3f" % (pd_i)
        self.report_data['pd_avg'] = self.data['pd_avg']
        # 表格：激光
        for i, ll_i in enumerate(self.data['ll']):
            self.report_data["ll_%d" % (i + 1)] = "%.3f" % (ll_i)
        for i, lr_i in enumerate(self.data['lr']):
            self.report_data["lr_%d" % (i + 1)] = "%.3f" % (lr_i)
        for i, ld_i in enumerate(self.data['ld']):
            self.report_data["ld_%d" % (i + 1)] = "%.3f" % (ld_i)
        self.report_data['ld_avg'] = self.data['ld_avg']
        # 数据处理
        self.report_data['lambda_y'] = "%.4f" % self.data['lambda_y']
        self.report_data['lambda_g'] = "%.4f" % self.data['lambda_g']
        self.report_data['lambda_p'] = "%.4f" % self.data['lambda_p']
        self.report_data['ua_d0'] = "%.5f" % self.data['ua_d0']
        self.report_data['u_d0'] = "%.5f" % self.data['u_d0']
        self.report_data['ua_dy'] = "%.5f" % self.data['ua_dy']
        self.report_data['u_dy'] = "%.5f" % self.data['u_dy']
        self.report_data['u_lambda_y'] = "%.3f" % self.data['u_lambda_y']
        self.report_data['ua_dg'] = "%.5f" % self.data['ua_dg']
        self.report_data['u_dg'] = "%.5f" % self.data['u_dg']
        self.report_data['u_lambda_g'] = "%.3f" % self.data['u_lambda_g']
        self.report_data['ua_dp'] = "%.5f" % self.data['ua_dp']
        self.report_data['u_dp'] = "%.5f" % self.data['u_dp']
        self.report_data['u_lambda_p'] = "%.3f" % self.data['u_lambda_p']
        # 调用ReportWriter类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    ll = Lloyd()
