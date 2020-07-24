import xlrd
from numpy import sqrt, abs

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method, Fitting
#from GeneralMethod.Report import Report


class Interferometry:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = [
        'l', 'b', 'h', 'lbd', 'm'
        '1', '2', '3', '4', '5', '6', '7', '8'
        'x-1', 'x-2', 'x-3', 'x-4', 'x-5', 'x-6', 'x-7', 'x-8'
        'y-1', 'y-2', 'y-3', 'y-4', 'y-5', 'y-6', 'y-7', 'y-8'
        'a', 'E', 'u_A', 'u_E', 'eta'
        'final'
    ]

    PREVIEW_FILENAME = "Preview.pdf"  # 预习报告模板文件的名称
    DATA_SHEET_FILENAME = "data.xlsx"  # 数据填写表格的名称
    REPORT_TEMPLATE_FILENAME = "Interferometry_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment2/2111Report.docx"  # 最后生成实验报告的相对路径

    def __init__(self):
        self.data = {}  # 存放实验中的各个物理量
        self.uncertainty = {}  # 存放物理量的不确定度
        self.report_data = {}  # 存放需要填入实验报告的
        print("2111 全息照相和全息干涉\n1. 实验预习\n2. 数据处理")
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
    从excel表格中读取数据，并创建填写key的list
    @param filename: 输入excel的文件名
    @return none
    '''
    def input_data(self, filename):
        ws = xlrd.open_workbook(filename).sheet_by_name('2111')
        list_x = []
        for col in range(1,9):
            list_x.append(float(ws.cell_value(3,col)) / 10)
        self.data['list_x'] = list_x
        self.data['num_l'] = (float(ws.cell_value(1,1)) / 100)
        self.data['num_b'] = (float(ws.cell_value(1,3)) / 100)
        self.data['num_h'] = (float(ws.cell_value(1,5)) / 100)
        self.data['num_lbd'] = (float(ws.cell_value(1,7)) / 1e9)
        self.data['num_m'] = (float(ws.cell_value(1,9)) / 1000)

    '''
    数据处理函数
    '''
    def calc_data(self):
        list_Y = []
        list_X = []
        list_x = self.data['list_x']
        for i, x in enumerate(list_x):
            list_Y.append((x ** 2) * (3 * self.data['num_l'] - x))
            list_X.append(2 * (i + 1) - 1)
        num_a, num_b, num_r = Fitting.linear(list_X, list_Y)
        num_para = 8 * self.data['num_m'] * 9.8 / (self.data['num_lbd'] * self.data['num_b'] * (self.data['num_h'] ** 3))
        num_E = num_a * num_para
        self.data['list_Y'] = list_Y
        self.data['list_X'] = list_X
        self.data['num_a'] = num_a
        self.data['num_r'] = num_r
        self.data['num_para'] = num_para
        self.data['num_E'] = num_E

    
    '''
    计算所有实验的不确定度
    '''
    def calc_uncertainty(self):
        num_sum = 0
        for i in range(0,8):
            num_sum += (self.data['list_Y'][i] - self.data['list_X'][i] * self.data['num_a']) ** 2
        num_u_A = sqrt(num_sum / (48 * 21))
        num_u_E = num_u_A * self.data['num_para']
        num_eta = abs(self.data['num_E']- 70 * 1e9) / (70 * 1e9) * 100
        self.data['num_u_A'] = num_u_A
        self.data['num_u_E'] = num_u_E
        self.data['num_eta'] = num_eta
        self.data['final'] = Method.final_expression(self.data['num_E'] / 1e9, self.data['num_u_E'] / 1e9)


    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 实验一
        # 表格：原始数据
        self.report_data['l'] = "%.2f" % (self.data['num_l'] * 100)
        self.report_data['b'] = "%.2f" % (self.data['num_b'] * 100)
        self.report_data['h'] = "%.2f" % (self.data['num_h'] * 100)
        self.report_data['lbd'] = "%.1f" % (self.data['num_lbd'] * 1e9)
        self.report_data['m'] = "%.2f" % (self.data['num_m'] * 1000)
        for i, x in enumerate(self.data['list_x']):
            self.report_data[str(i + 1)] = "%.2f" % (x * 10)
        # 最终结果
        self.report_data['final'] = self.data['final']
        # 将各个变量以及不确定度的结果导入实验报告，在实际编写中需根据实验报告的具体要求设定保留几位小数
        for i, x in enumerate(self.data['list_X']):
            self.report_data['x-'+str(i + 1)] = "%.12f" % x
        for i, y in enumerate(self.data['list_Y']):
            self.report_data['y-'+str(i + 1)] = "%.12f" % y
        self.report_data['a'] = "%.12f" % self.data['num_a']
        self.report_data['E'] = "%.2f" % (self.data['num_E'] / 1e9)
        self.report_data['u_A'] = "%.12f" % self.data['num_u_A']
        self.report_data['u_E'] = "%.12f" % (self.data['num_u_E'] / 1e9)
        self.report_data['eta'] = "%.1f" % self.data['num_eta']

        # 调用ReportWriter类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)


if __name__ == '__main__':
    itf = Interferometry()
