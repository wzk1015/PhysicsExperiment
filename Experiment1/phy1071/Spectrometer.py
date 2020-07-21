import xlrd
import shutil
import os
from numpy import sqrt, abs, sin, cos, pi
import subprocess

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method
from GeneralMethod.Report import Report


class Spectrometer:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = []

    PREVIEW_FILENAME = "Preview.pdf"  # 预习报告模板文件的名称
    DATA_SHEET_FILENAME = "data.xlsx"  # 数据填写表格的名称
    REPORT_TEMPLATE_FILENAME = "Spectrometer_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment1/1071Report.docx"  # 最后生成实验报告的相对路径

    def __init__(self):
        self.data = {}  # 存放实验中的各个物理量
        self.uncertainty = {}  # 存放物理量的不确定度
        self.report_data = {}  # 存放需要填入实验报告的
        print("1071 分光仪调整及其应用实验\n1. 实验预习\n2. 数据处理")
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
        ws = xlrd.open_workbook(filename).sheet_by_name('Sheet1')
        # 从excel中读取数据
        list_data = []  # 实验数据：0是实验一，1是实验二，2是实验三
        i = 0
        for col in [1, 7, 13]:
            list_data.append([])
            for row in range(3, 12):
                line = ws.row_values(row, col, col+4)
                if line[0] != '':
                    list_data[i].append(line)    # 从excel取出来的数据
            i += 1
        self.data['list_data'] = list_data  # 存储从表格中读入的数据

    '''
    自定义角度转换方法，考虑是否需要加入到通用工具
    @param
    degree: 角度的字符串表示，度和分之间以空格隔开
    @return
    degree_value: float形式的角度，单位为度
    '''
    @staticmethod
    def degree_trans(degree):
        degree_list = degree.split(' ')
        if len(degree_list) == 1:
            return eval(degree)
        elif len(degree_list) == 2:
            return float(eval(degree_list[0])) + float(eval(degree_list[1]) / 60.0)
        else:
            return False

    '''
    数据处理总函数，调用三个实验的数据处理函数
    '''
    def calc_data(self):
        self.calc_data1()
        self.calc_data2()
        self.calc_data3()

    '''
    进行实验一数据处理
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    def calc_data1(self):
        list_A = []
        for rows in self.data['list_data'][0]:
            row = []
            for item in rows:
                value_degree = self.degree_trans(item)
                if (value_degree > 1000) | (value_degree < -1000):
                    print('数据出现错误，请检查实验表格第'
                          + str(self.data['list_data'].index(rows) + 4) + '行第' + str(rows.index(item) + 2)
                          + '列的度和分是否以空格分隔')
                else:
                    row.append(value_degree)
            value_A = (row[3] + row[2] - row[1] - row[0]) / 4
            if value_A < 0:
                value_A = value_A + 90
            list_A.append(value_A)
        aver_A = Method.average(list_A)
        self.data['list_A'] = list_A
        self.data['aver_A'] = aver_A

    '''
    进行实验二数据处理
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    def calc_data2(self):
        list_delta_m = []
        for rows in self.data['list_data'][1]:
            row = []
            for item in rows:
                value_degree = self.degree_trans(item)
                if (value_degree > 1000) | (value_degree < -1000):
                    print('数据出现错误，请检查实验表格第'
                          + str(self.data['list_data'].index(rows) + 4) + '行第' + str(rows.index(item) + 8)
                          + '列的度和分是否以空格分隔')
                else:
                    row.append(value_degree)
            value_delta_m = (row[3] + row[2] - row[1] - row[0]) / 2
            if value_delta_m < 0:
                value_delta_m = value_delta_m + 180
            list_delta_m.append(value_delta_m)
        delta_m = Method.average(list_delta_m)
        n1 = float(sin((delta_m + self.data['aver_A']) / 360 * pi) / sin(self.data['aver_A'] / 360 * pi))
        self.data['list_delta_m'] = list_delta_m
        self.data['aver_delta_m'] = delta_m
        self.data['n1'] = n1

    '''
    进行实验三数据处理
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    def calc_data3(self):
        list_delta = []
        for rows in self.data['list_data'][2]:
            row = []
            for item in rows:
                value_degree = self.degree_trans(item)
                if (value_degree > 1000) | (value_degree < -1000):
                    print('数据出现错误，请检查实验表格第'
                          + str(self.data['list_data'].index(rows) + 4) + '行第' + str(rows.index(item) + 14)
                          + '列的度和分是否以空格分隔')
                else:
                    row.append(value_degree)
            value_delta = (row[0] + row[1] - row[2] - row[3]) / 2
            if value_delta < 0:
                value_delta = value_delta + 180
            list_delta.append(value_delta)
        delta = Method.average(list_delta)
        n2 = float(sqrt(((cos(self.data['aver_A'] / 180 * pi) + sin(delta / 180 * pi)) / sin(self.data['aver_A'] / 180 * pi)) ** 2 + 1))
        self.data['list_delta'] = list_delta
        self.data['aver_delta'] = delta
        self.data['n2'] = n2

    '''
    计算所有实验的不确定度
    '''
    def calc_uncertainty(self):
        self.calc_uncertainty1()
        self.calc_uncertainty2()
        self.calc_uncertainty3()

    '''
    计算实验一的不确定度
    '''
    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    def calc_uncertainty1(self):
        # 计算顶角A的a,b及总不确定度
        ua_A = Method.a_uncertainty(self.data['list_A']) # 这里容易写错，一定要用原始数据的数组
        ub_A = (1 / 60) / sqrt(3) / 2
        u_A = sqrt(ua_A ** 2 + ub_A ** 2)
        self.data.update({"ua_A": ua_A, "ub_A": ub_A, "u_A": u_A})
        self.data['final_A'] = Method.final_expression(self.data['aver_A'], u_A)

    '''
    计算实验二的不确定度
    '''
    def calc_uncertainty2(self):
        # 计算顶角A的a,b及总不确定度
        ua_delta_m = Method.a_uncertainty(self.data['list_delta_m'])  # 这里容易写错，一定要用原始数据的数组
        ub_delta_m = (1 / 60) / sqrt(3)
        u_delta_m = sqrt(ua_delta_m ** 2 + ub_delta_m ** 2)
        self.data.update({"ua_delta_m": ua_delta_m, "ub_delta_m": ub_delta_m, "u_delta_m": u_delta_m})
        part_1 = cos((self.data['aver_delta_m'] / 180 * pi + self.data['aver_A'] / 180 * pi) / 2)
        part_4 = sin(self.data['aver_A'] / 180 * pi / 2)
        n_delta_m = 0.5 * part_1 / part_4
        n_A = -sin(self.data['aver_delta_m'] / 180 * pi / 2) / 2 / (part_4 ** 2)
        u_n1 = sqrt((n_delta_m * u_delta_m / 180 * pi) ** 2 + (n_A * self.data['u_A'] / 180 * pi) ** 2)
        self.data['u_n1'] = u_n1
        self.data['final_n1'] = Method.final_expression(self.data['n1'], u_n1)

    '''
    计算实验三的不确定度
    '''
    def calc_uncertainty3(self):
        # 计算顶角A的a,b及总不确定度
        ua_delta = Method.a_uncertainty(self.data['list_delta'])  # 这里容易写错，一定要用原始数据的数组
        ub_delta = (1 / 60) / sqrt(3)
        u_delta = sqrt(ua_delta ** 2 + ub_delta ** 2)
        self.data.update({"ua_delta": ua_delta, "ub_delta": ub_delta, "u_delta": u_delta})
        part_1 = - 1 / sin(self.data['aver_A'] / 180 * pi) - sin(self.data['aver_delta'] / 180 * pi) * cos(self.data['aver_A'] / 180 * pi) / (sin(self.data['aver_A'] / 180 * pi) ** 2)
        part_4 = (cos(self.data['aver_A'] / 180 * pi) + sin(self.data['aver_delta'] / 180 * pi)) / sin(self.data['aver_A'] / 180 * pi)
        n_delta = cos(self.data['aver_delta'] / 180 * pi) * part_4 / self.data['n2']
        n_A = part_1 * part_4 / self.data['n2']
        print(n_delta)
        print(n_A)
        u_n2 = sqrt((n_delta * u_delta / 180 * pi) ** 2 + (n_A * self.data['u_A'] / 180 * pi) ** 2)
        self.data['u_n2'] = u_n2
        self.data['final_n2'] = Method.final_expression(self.data['n2'], u_n2)

    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 表格：实验一原始数据
        for k, data_list in enumerate(self.data['list_data']):
            for i, degree_row_i in enumerate(data_list):
                for j, degree_j in enumerate(degree_row_i):
                    degree_list = degree_j.split(' ')
                    if j == 0:
                        self.report_data['s' + str(k+1) + '1' + str(i + 1) + 'a1'] = \
                            degree_list[0] + '°' + degree_list[1] + '\''  # 一定都是字符串类型
                    elif j == 1:
                        self.report_data['s' + str(k+1) + '1' + str(i + 1) + 'b1'] = \
                            degree_list[0] + '°' + degree_list[1] + '\''
                    elif j == 2:
                        self.report_data['s' + str(k+1) + '1' + str(i + 1) + 'a2'] = \
                            degree_list[0] + '°' + degree_list[1] + '\''
                    else:
                        self.report_data['s' + str(k+1) + '1' + str(i + 1) + 'b2'] = \
                            degree_list[0] + '°' + degree_list[1] + '\''
            for i in range(len(data_list), 10):
                self.report_data['s' + str(k + 1) + '1' + str(i + 1) + 'a1'] = ''
                self.report_data['s' + str(k + 1) + '1' + str(i + 1) + 'a2'] = ''
                self.report_data['s' + str(k + 1) + '1' + str(i + 1) + 'b1'] = ''
                self.report_data['s' + str(k + 1) + '1' + str(i + 1) + 'b2'] = ''
        # 表格：三个计算值
        for i, A in enumerate(self.data['list_A']):
            self.report_data['s12' + str(i + 1)] = "%.3f" % A
        for i in range(len(self.data['list_A']), 10):
            self.report_data['s12' + str(i + 1)] = ''
        for i, delta_m in enumerate(self.data['list_delta_m']):
            self.report_data['s22' + str(i+1)] = "%.3f" % delta_m
        for i in range(len(self.data['list_delta_m']), 10):
            self.report_data['s22' + str(i + 1)] = ''
        for i, delta in enumerate(self.data['list_delta']):
            self.report_data['s32' + str(i+1)] = "%.3f" % delta
        for i in range(len(self.data['list_delta']), 10):
            self.report_data['s32' + str(i + 1)] = ''
        # 最终结果
        self.report_data['s12final_A'] = self.data['final_A']
        self.report_data['s22final_n'] = self.data['final_n1']
        self.report_data['s32final_n'] = self.data['final_n2']
        # 将各个变量以及不确定度的结果导入实验报告，在实际编写中需根据实验报告的具体要求设定保留几位小数
        self.report_data['s12calcu_A'] = "%.3f" % self.data['aver_A']
        self.report_data['s12ua_A'] = "%.5f" % self.data['ua_A']
        self.report_data['s12ub_A'] = "%.5f" % self.data['ub_A']
        self.report_data['s12u_A'] = "%.5f" % self.data['u_A']
        self.report_data['s22calcu_delta_m'] = "%.3f" % self.data['aver_delta_m']
        self.report_data['s22calcu_n'] = "%.3f" % self.data['n1']
        self.report_data['s22ua_delta_m'] = "%.5f" % self.data['ua_delta_m']
        self.report_data['s22ub_delta_m'] = "%.5f" % self.data['ub_delta_m']
        self.report_data['s22u_delta_m'] = "%.5f" % self.data['u_delta_m']
        self.report_data['s22u_n'] = "%.5f" % self.data['u_n1']
        self.report_data['s32calcu_delta'] = "%.3f" % self.data['aver_delta']
        self.report_data['s32calcu_n'] = "%.3f" % self.data['n2']
        self.report_data['s32ua_delta'] = "%.5f" % self.data['ua_delta']
        self.report_data['s32ub_delta'] = "%.5f" % self.data['ub_delta']
        self.report_data['s32u_delta'] = "%.5f" % self.data['u_delta']
        self.report_data['s32u_n'] = "%.5f" % self.data['u_n2']
        # 调用ReportWriter类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)


if __name__ == '__main__':
    mc = Spectrometer()
