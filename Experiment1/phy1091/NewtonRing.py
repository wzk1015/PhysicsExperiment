import xlrd
import os, sys, shutil
from numpy import asarray, sqrt
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method, Fitting
from GeneralMethod.Report import Report


class NewtonRing :
    PREVIEW_FILENAME = "Preview.pdf"  # 预习报告模板文件的名称
    DATA_SHEET_FILENAME = "data.xlsx"  # 数据填写表格的名称
    REPORT_TEMPLATE_FILENAME = "Newton_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment1/1091Report_2.docx"  # 最后生成实验报告的相对路径

    def __init__(self):
        self.data = {}
        self.uncertainty = {}
        self.report_data = {}
        print("1091-2 牛顿环干涉实验")
        print("1. 实验预习\n2. 数据处理")
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
        pass
    

    def input_data(self, filename):
        ws = xlrd.open_workbook(self.DATA_SHEET_FILENAME).sheet_by_name('Newton')
        arr_k = []
        arr_x1 = []
        arr_x2 = []
        arr_D = []
        arr_D_sq = []
        for col in range(1, 11):
            arr_k.append(int(ws.cell_value(1, col)))
            arr_x1.append(float(ws.cell_value(2, col)))
            arr_x2.append(float(ws.cell_value(3, col)))
            arr_D.append(float(ws.cell_value(4, col)))
            arr_D_sq.append(float(ws.cell_value(5, col)))
        self.data['arr_k'] = arr_k
        self.data['arr_x1'] = arr_x1
        self.data['arr_x2'] = arr_x2
        self.data['arr_D'] = arr_D
        self.data['arr_D_sq'] = arr_D_sq

        self.data['lambda'] = float(ws.cell_value(7, 3))

        # finished

    def calc_data(self):
        x = arr_k = asarray(self.data['arr_k'])
        y = arr_D_sq = asarray(self.data['arr_D_sq'])
        b, a, r = Fitting.linear(x, y)
        self.data['b'] = b
        self.data['a'] = a
        self.data['r'] = r
        x_mean = Method.average(x)
        y_mean = Method.average(y)
        x_sq_mean = Method.average(x ** 2)
        y_sq_mean = Method.average(y ** 2)
        xy_mean = Method.average(x * y)

        self.data['x_mean'] = x_mean
        self.data['y_mean'] = y_mean
        self.data['x_sq_mean'] = x_sq_mean
        self.data['y_sq_mean'] = y_sq_mean
        self.data['xy_mean'] = xy_mean

        lbd = self.data['lambda'] * 1e-6
        R = b / (4 * lbd)
        self.data['R'] = R
        pass

    def calc_uncertainty(self):
        b = self.data['b']
        n = len(self.data['arr_D'])
        r = self.data['r']
        lbd = self.data['lambda'] * 1e-6
        ua_b = b * sqrt((1 / (n - 2)) * ((1 / (r ** 2)) - 1))
        ub_b = 0.005 / sqrt(3)
        u_b = sqrt(ua_b ** 2 + ub_b ** 2)
        u_R = u_b / (4 * lbd)
        R = self.data['R']
        res, unc, pwr = Method.final_expression(R, u_R)
        self.uncertainty['ua_b'] = ua_b
        self.uncertainty['ub_b'] = ub_b
        self.uncertainty['u_b'] = u_b
        self.uncertainty['u_R'] = u_R
        self.data['final'] = "(%.0f ± %.0f)e%d" % (res, unc, int(pwr))
        pass

    def fill_report(self):
        # 先填表
        arr_k = self.data['arr_k']
        arr_x1 = self.data['arr_x1']
        arr_x2 = self.data['arr_x2']
        arr_D = self.data['arr_D']
        arr_D_sq = self.data['arr_D_sq']
        for i, k_i in enumerate(arr_k):
            kw = "k%d" % (i + 1)
            self.report_data[kw] = "%d" % k_i
        for i, x1_i in enumerate(arr_x1):
            kw = "x1-%d" % (i + 1)
            self.report_data[kw] = "%g" % (x1_i)
        for i, x2_i in enumerate(arr_x2):
            kw = "x2-%d" % (i + 1)
            self.report_data[kw] = "%g" % (x2_i)
        for i, D_i in enumerate(arr_D):
            kw = "D-%d" % (i + 1)
            self.report_data[kw] = "%g" % (D_i)
        for i, D_sq_i in enumerate(arr_D_sq):
            kw = "D2-%d" % (i + 1)
            self.report_data[kw] = "%g" % (D_sq_i)
        self.report_data['x-mean'] = "%g" % self.data['x_mean']
        self.report_data['y-mean'] = "%g" % self.data['y_mean']
        self.report_data['x2-mean'] = "%g" % self.data['x_sq_mean']
        self.report_data['y2-mean'] = "%g" % self.data['y_sq_mean']
        self.report_data['xy-mean'] = "%g" % self.data['xy_mean']
        self.report_data['a'] = "%g" % self.data['a']
        self.report_data['b'] = "%g" % self.data['b']
        self.report_data['r'] = "%g" % self.data['r']

        self.report_data['lbd'] = "%g" % self.data['lambda']
        self.report_data['res-R'] = "%g" % self.data['R']

        self.report_data['ua-b'] = "%g" % self.uncertainty['ua_b']
        self.report_data['ub-b'] = "%g" % self.uncertainty['ub_b']
        self.report_data['u-b'] = "%g" % self.uncertainty['u_b']
        self.report_data['u-R'] = "%g" % self.uncertainty['u_R']

        self.report_data['final'] = self.data['final']
        
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

    pass

if __name__ == '__main__':
    NewtonRing()