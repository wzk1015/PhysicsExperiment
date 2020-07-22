import xlrd
from numpy import sqrt, abs, sin, cos, pi

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method, Simplified
from GeneralMethod.Report import Report


class Spectrometer:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = []

    PREVIEW_FILENAME = "Preview.pdf"  # 预习报告模板文件的名称
    DATA_SHEET_FILENAME = "data.xlsx"  # 数据填写表格的名称
    REPORT_TEMPLATE_FILENAME = "Laser_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment1/1081Report.docx"  # 最后生成实验报告的相对路径

    def __init__(self):
        self.data = {}  # 存放实验中的各个物理量
        self.uncertainty = {}  # 存放物理量的不确定度
        self.report_data = {}  # 存放需要填入实验报告的
        print("1081 光的干涉实验\n1. 实验预习\n2. 数据处理")
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
        ws_biprism = xlrd.open_workbook(filename).sheet_by_name('Biprism')
        ws_lloyd = xlrd.open_workbook(filename).sheet_by_name('Lloyd')
        # 从excel中读取数据
        list_data = []  # 实验数据：0是实验一，1是实验二
        add_item = {}
        exp_data = []
        for row in [2, 4, 6, 8]:
            for col in range(1, 6):
                exp_data.append(float(ws_biprism.cell_value(row, col)))
        add_item['exp_data'] = exp_data
        add_item['K'] = float(ws_biprism.cell_value(13, 1))
        add_item['B'] = float(ws_biprism.cell_value(13, 2))
        add_item['L1'] = float(ws_biprism.cell_value(13, 3))
        add_item['L2'] = float(ws_biprism.cell_value(13, 4))
        add_item['E'] = float(ws_biprism.cell_value(13, 5))
        add_item['bL'] = float(ws_biprism.cell_value(17, 2))
        add_item['bR'] = float(ws_biprism.cell_value(17, 3))
        add_item['b1L'] = float(ws_biprism.cell_value(18, 2))
        add_item['b1R'] = float(ws_biprism.cell_value(18, 3))
        add_item['lambda_0'] = float(ws_biprism.cell_value(1, 8))
        add_item['delta_b_b'] = float(ws_biprism.cell_value(2, 8))
        add_item['delta_s'] = float(ws_biprism.cell_value(3, 8))
        list_data.append(add_item)
        add_item = {}
        exp_data = []
        for row in [2, 4, 6, 8]:
            for col in range(1, 6):
                exp_data.append(float(ws_lloyd.cell_value(row, col)))
        add_item['exp_data'] = exp_data
        add_item['K'] = float(ws_lloyd.cell_value(13, 1))
        add_item['B'] = float(ws_lloyd.cell_value(13, 2))
        add_item['L1'] = float(ws_lloyd.cell_value(13, 3))
        add_item['L2'] = float(ws_lloyd.cell_value(13, 4))
        add_item['E'] = float(ws_lloyd.cell_value(13, 5))
        add_item['bL'] = float(ws_lloyd.cell_value(17, 2))
        add_item['bR'] = float(ws_lloyd.cell_value(17, 3))
        add_item['b1L'] = float(ws_lloyd.cell_value(18, 2))
        add_item['b1R'] = float(ws_lloyd.cell_value(18, 3))
        add_item['lambda_0'] = float(ws_lloyd.cell_value(1, 8))
        add_item['delta_b_b'] = float(ws_lloyd.cell_value(2, 8))
        add_item['delta_s'] = float(ws_lloyd.cell_value(3, 8))
        list_data.append(add_item)
        self.data['list_data'] = list_data  # 存储从表格中读入的数据

    '''
    数据处理总函数，调用三个实验的数据处理函数
    '''
    def calc_data(self):
        result_list = []
        add_result = self.calc_data_detail(self.data['list_data'][0])
        result_list.append(add_result)
        add_result = self.calc_data_detail(self.data['list_data'][1])
        result_list.append(add_result)
        self.data['result_data'] = result_list

    '''
    具体实验数据处理，因为两个实验除了实验数据以外都相同，因此可以使用同一函数。
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    @staticmethod
    def calc_data_detail(data_list):
        result_data = {}
        diff_10delta_x, aver_10delta_x = Method.successive_diff(data_list['exp_data'])
        delta_x = aver_10delta_x / 10
        b = abs(data_list['bL'] - data_list['bR'])
        b1 = abs(data_list['b1L'] - data_list['b1R'])
        a = sqrt(b * b1)
        S = abs(data_list['K'] - data_list['L2'])
        S1 = abs(data_list['K'] - data_list['L1'])
        D = S + S1
        result_lambda = a / D * delta_x * 100000
        error = abs(result_lambda - data_list['lambda_0']) / data_list['lambda_0'] * 100
        result_data['diff_10delta_x'] = diff_10delta_x
        result_data['aver_10delta_x'] = aver_10delta_x
        result_data['delta_x'] = delta_x
        result_data['b'] = b
        result_data['b1'] = b1
        result_data['a'] = a
        result_data['S'] = S
        result_data['S1'] = S1
        result_data['D'] = D
        result_data['result_lambda'] = result_lambda
        result_data['error'] = error
        return result_data

    '''
    计算所有实验的不确定度
    '''
    def calc_uncertainty(self):
        uncertain_list = []
        add_uncertain = self.calc_uncertainty_detail(self.data['list_data'][0], self.data['result_data'][0])
        uncertain_list.append(add_uncertain)
        add_uncertain = self.calc_uncertainty_detail(self.data['list_data'][1], self.data['result_data'][1])
        uncertain_list.append(add_uncertain)
        self.data['uncertain_data'] = uncertain_list


    '''
    具体实验的不确定度，两个实验相同
    '''
    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    @staticmethod
    def calc_uncertainty_detail(data_list, result_list):
        sunc = Simplified()
        uncertain_data = {}
        ua_10delta_x = Method.a_uncertainty(result_list['diff_10delta_x'])
        ub_10delta_x = sunc.micrometer / sqrt(3)
        u_10delta_x = sqrt(ua_10delta_x ** 2 + ub_10delta_x ** 2)
        u_delta_x = u_10delta_x / 10
        ub_S = sqrt((sunc.steel_ruler / 10 / sqrt(3)) ** 2 + (data_list['delta_s'] / sqrt(3)) ** 2)
        ub_b = sqrt(
            (sunc.micrometer / sqrt(3)) ** 2 + (data_list['delta_b_b'] * result_list['b'] / sqrt(3)) ** 2)
        ub_b1 = sqrt(
            (sunc.micrometer / sqrt(3)) ** 2 + (data_list['delta_b_b'] * result_list['b1'] / sqrt(3)) ** 2)
        u_lbd_lbd = sqrt(
            (u_delta_x / result_list['delta_x']) ** 2 +
            (ub_b / result_list['b'] / 2) ** 2 +
            (ub_b1 / result_list['b1'] / 2) ** 2 +
            2 * ((ub_S / (result_list['S'] + result_list['S1'])) ** 2)
        )
        u_lbd = u_lbd_lbd * result_list['result_lambda']
        final_lbd = Method.final_expression(result_list['result_lambda'], u_lbd)
        uncertain_data['ua_10delta_x'] = ua_10delta_x
        uncertain_data['ub_10delta_x'] = ub_10delta_x
        uncertain_data['u_10delta_x'] = u_10delta_x
        uncertain_data['u_delta_x'] = u_delta_x
        uncertain_data['u_S'] = ub_S
        uncertain_data['u_b'] = ub_b
        uncertain_data['u_b1'] = ub_b1
        uncertain_data['u_lbd_lbd'] = u_lbd_lbd
        uncertain_data['u_lbd'] = u_lbd
        uncertain_data['final'] = final_lbd
        return uncertain_data

    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        for exp in range(2):
            # 表格：实验原始数据
            for i, x in enumerate(self.data['list_data'][exp]['exp_data']):
                self.report_data['s' + str(exp + 1) + '1' + str(i + 1)] = "%.3f" % x
            self.report_data['s' + str(exp + 1) + '1K'] = "%.2f" % self.data['list_data'][exp]['K']
            self.report_data['s' + str(exp + 1) + '1B'] = "%.2f" % self.data['list_data'][exp]['B']
            self.report_data['s' + str(exp + 1) + '1L1'] = "%.2f" % self.data['list_data'][exp]['L1']
            self.report_data['s' + str(exp + 1) + '1L2'] = "%.2f" % self.data['list_data'][exp]['L2']
            self.report_data['s' + str(exp + 1) + '1E'] = "%.2f" % self.data['list_data'][exp]['E']
            self.report_data['s' + str(exp + 1) + '1bL'] = "%.3f" % self.data['list_data'][exp]['bL']
            self.report_data['s' + str(exp + 1) + '1bR'] = "%.3f" % self.data['list_data'][exp]['bR']
            #  替换的时候大小写不分？真的想骂人了
            self.report_data['s' + str(exp + 1) + '1sb'] = "%.3f" % self.data['result_data'][exp]['b']
            self.report_data['s' + str(exp + 1) + '1b1L'] = "%.3f" % self.data['list_data'][exp]['b1L']
            self.report_data['s' + str(exp + 1) + '1b1R'] = "%.3f" % self.data['list_data'][exp]['b1R']
            self.report_data['s' + str(exp + 1) + '1b1'] = "%.3f" % self.data['result_data'][exp]['b1']
            # 表格：逐差法计算值
            for i, diff in enumerate(self.data['result_data'][exp]['diff_10delta_x']):
                self.report_data['s' + str(exp + 1) + '2' + str(i + 1)] = "%.3f" % diff
            # 最终结果
            self.report_data['s' + str(exp + 1) + '2final_lambda'] = self.data['uncertain_data'][exp]['final']
            # 将各个变量以及不确定度的结果导入实验报告，在实际编写中需根据实验报告的具体要求设定保留几位小数
            self.report_data['s' + str(exp + 1) + '210_delta_x'] = "%.3f" % self.data['result_data'][exp]['aver_10delta_x']
            self.report_data['s' + str(exp + 1) + '2_delta_x'] = "%.3f" % self.data['result_data'][exp]['delta_x']
            self.report_data['s' + str(exp + 1) + '2_a'] = "%.3f" % self.data['result_data'][exp]['a']
            self.report_data['s' + str(exp + 1) + '2_S'] = "%.2f" % self.data['result_data'][exp]['S']
            self.report_data['s' + str(exp + 1) + '2_S1'] = "%.2f" % self.data['result_data'][exp]['S1']
            self.report_data['s' + str(exp + 1) + '2_D'] = "%.2f" % self.data['result_data'][exp]['D']
            self.report_data['s' + str(exp + 1) + '2_lambda'] = "%.2f" % self.data['result_data'][exp]['result_lambda']
            self.report_data['s' + str(exp + 1) + '2_lambda0'] = "%.2f" % self.data['list_data'][exp]['lambda_0']
            self.report_data['s' + str(exp + 1) + '2_error'] = "%.5f" % self.data['result_data'][exp]['error']
            self.report_data['s' + str(exp + 1) + '2_db'] = "%.2f" % self.data['list_data'][exp]['delta_b_b']
            self.report_data['s' + str(exp + 1) + '2_dS'] = "%.1f" % self.data['list_data'][exp]['delta_s']
            self.report_data['s' + str(exp + 1) + '2ua_10dx'] = "%.3f" % self.data['uncertain_data'][exp]['ua_10delta_x']
            self.report_data['s' + str(exp + 1) + '2ub_10dx'] = "%.3f" % self.data['uncertain_data'][exp]['ub_10delta_x']
            self.report_data['s' + str(exp + 1) + '2u_10dx'] = "%.3f" % self.data['uncertain_data'][exp]['u_10delta_x']
            self.report_data['s' + str(exp + 1) + '2u_dx'] = "%.3f" % self.data['uncertain_data'][exp]['u_delta_x']
            self.report_data['s' + str(exp + 1) + '2u_S'] = "%.2f" % self.data['uncertain_data'][exp]['u_S']
            self.report_data['s' + str(exp + 1) + '2u_b'] = "%.3f" % self.data['uncertain_data'][exp]['u_b']
            self.report_data['s' + str(exp + 1) + '2u_b1'] = "%.3f" % self.data['uncertain_data'][exp]['u_b1']
            self.report_data['s' + str(exp + 1) + '2u_lbd_lbd'] = "%.3f" % self.data['uncertain_data'][exp]['u_lbd_lbd']
            self.report_data['s' + str(exp + 1) + '2u_lambda'] = "%.3f" % self.data['uncertain_data'][exp]['u_lbd']
        # 调用ReportWriter类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)


if __name__ == '__main__':
    mc = Spectrometer()
