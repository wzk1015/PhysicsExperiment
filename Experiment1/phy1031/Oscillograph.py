import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
from numpy import sqrt, abs

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method,Fitting
from reportwriter.ReportWriter import ReportWriter



class Oscillograph:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = [
        "1", "2", "3", "4", "5", "6", "7", "8", "9", "10",
        "11", "12", "13", "14", "15", "16", "17", "18", "19", "20",
        "10d-1", "10d-2","10d-3","10d-4","10d-5","10d-6","10d-7","10d-8","10d-9","10d-10",# 逐差法；10Δd
        "d", "lbd","f_1","f_2",
        "ua_d", "ub1_d","ub2_d", "u_d","u_lbd","fin_lbd",
        "ave_f","d_f","u_f","fin_f",
        "c","u_c_c","u_c","fin_c"
    ]

    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "data.xlsx"
    REPORT_TEMPLATE_FILENAME = "Oscillograph_empty.docx"
    REPORT_OUTPUT_FILENAME = "Oscillograph_out.docx"

    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {} # 存放需要填入实验报告的
        print("1031 示波器的使用\n1. 实验预习\n2. 数据处理")
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
            # 计算不确定度
            self.calc_uncertainty()
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            os.startfile(self.REPORT_OUTPUT_FILENAME)
            print("Done!")

    def input_data(self, filename):
        ws = xlrd.open_workbook(filename).sheet_by_name('Oscillograph')
        # 从excel中读取数据
        list_d1 = []
        list_d2 = []
        list_f = []
        for col in range(1, 11):
            list_d1.append(float(ws.cell_value(2, col)))
            list_d2.append(float(ws.cell_value(4, col)))
        list_f.append(float(ws.cell_value(7, 2)))
        list_f.append(float(ws.cell_value(7, 4)))


        self.data['list_d1'] = list_d1
        self.data['list_d2'] = list_d2
         # 存储从表格中读入的数据
        self.data['list_f'] = list_f
    '''
    进行数据处理
    '''
    def calc_data(self):
        # list_dif_d 长度为x折半的数组，为逐差相减的结果
        # num_d 逐差法求得的平均值
        list_dtx = []
        for i in range(0, 10):
            list_dtx.append((self.data['list_d2'][i]-self.data['list_d1'][i]) / 30)
        num_d = Method.average(list_dtx)
        num_dt_l = 30 * num_d
        num_lbd = 2 * num_d
        self.data['list_dtx'] = list_dtx
        self.data['num_d'] = num_d
        # num_lbd = round(num_lbd,5)
        num_ave_f= (float((self.data['list_f'][0]+self.data['list_f'][1])/2.0))
        num_delta_f = abs(self.data['list_f'][0]-self.data['list_f'][1])/2.0
        num_f = num_ave_f
        num_c = num_lbd * num_f
        self.data['num_lbd'] = num_lbd
        self.data['num_ave_f'] = num_ave_f
        self.data['num_delta_f'] = num_delta_f
        self.data['num_u_f'] = num_delta_f/sqrt(3)
        self.data['num_c'] = num_c
        self.data['num_f'] = num_f
 
    '''
    计算所有的不确定度
    '''
    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    def calc_uncertainty(self):
        list_tmp = []
        for i in self.data['list_dtx']:
            list_tmp.append(i * 0.001 / 15)
        # num_ua_d = sqrt( (Method.variance(self.data['list_d'])) / (10 * 9) )
        num_ua_d = Method.a_uncertainty(list_tmp) 
        # print(num_ua_d)
        num_ub1_d = 0.005 / sqrt(3)
        num_ub2_d = 0.1 / sqrt(3)
        num_u_d = sqrt(num_ua_d ** 2 + num_ub1_d ** 2 + num_ub2_d ** 2)
        self.data.update({"num_ua_d":num_ua_d, "num_ub1_d":num_ub1_d, "num_ub2_d":num_ub2_d, "num_u_d":num_u_d})
        num_d = self.data['num_d']
        num_lbd = self.data['num_lbd']
        num_c = self.data['num_c']
        num_f = self.data['num_f']
        num_u_f = self.data['num_u_f']
        num_u_lbd = 2 * num_u_d
        num_u_c_c = sqrt( (num_u_lbd/num_lbd) ** 2 + (num_u_f/num_f) ** 2)
        num_u_c = num_u_c_c * num_c
        self.data['num_u_c'] = num_u_c
        self.data['num_u_c_c']=num_u_c_c
        self.data.update({"num_u_lbd": num_u_lbd})
        num_u_lbd_1bit, pwr = Method.scientific_notation(num_u_lbd)
        num_u_f_1bit, pwr = Method.scientific_notation(num_u_f) # 将不确定度转化为只有一位有效数字的科学计数法
        num_u_c_1bit, pwr = Method.scientific_notation(num_u_c) # 将不确定度转化为只有一位有效数字的科学计数法
        num_fin_lbd = int(num_lbd * (10 ** pwr)) / (10 ** pwr) # 对物理量保留有效数字，截断处理
        num_fin_f = int(num_f * (10 ** pwr)) / (10 ** pwr) 
        num_fin_c = int(num_c * (10 ** pwr)) / (10 ** pwr) 
        self.data['num_fin_c'] = "%.2f±%.2f" % (num_fin_c,num_u_c_1bit)
        self.data['num_fin_f'] = "%.0f±%.0f" % (num_fin_f,num_u_f_1bit)
        self.data['num_fin_lbd'] = "%.0f±%.0f" % (num_fin_lbd,num_u_lbd_1bit)
    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 表格：原始数据d
        for i, d_i in enumerate(self.data['list_d1']):
            self.report_data[str(i + 1)] = "%.5f" % (d_i) # 一定都是字符串类型
        # 表格：逐差法计算10Δd
        for i, d_i in enumerate(self.data['list_d2']):
            self.report_data[str(i + 10)] = "%.5f" % (d_i)
        for i, dif_d_i in enumerate(self.data['list_dtx']):
            self.report_data["10d-%d" % (i + 1)] = "%.5f" % (dif_d_i)
        # 最终结果
        # 将各个变量以及不确定度的结果导入实验报告，在实际编写中需根据实验报告的具体要求设定保留几位小数
        self.report_data['f_1'] = "%.4f" % self.data['list_f'][0]
        self.report_data['f_2'] = "%.4f" % self.data['list_f'][1]
        self.report_data['ave_f'] = "%.4f" % self.data['num_ave_f']
        self.report_data['d_f'] = "%.4f" % self.data['num_delta_f']
        self.report_data['fin_lbd'] = self.data['num_fin_lbd']
        self.report_data['u_f'] = "%.5f" % self.data['num_u_f']
        self.report_data['c'] = "%.5f" % self.data['num_c']
        self.report_data['u_c_c'] = "%.4f" % self.data['num_u_c_c']
        self.report_data['u_c'] = "%.4f" % self.data['num_u_c']
        self.report_data['fin_c'] = self.data['num_fin_c']
        self.report_data['fin_f'] = self.data['num_fin_f']
        self.report_data['d'] = "%.5f" % self.data['num_d']
        self.report_data['lbd'] = "%.5f" % self.data['num_lbd']
        self.report_data['ua_d'] = "%.4f" % self.data['num_ua_d']
        self.report_data['ub1_d'] = "%.4f" % self.data['num_ub1_d']
        self.report_data['ub2_d'] = "%.4f" % self.data['num_ub2_d']
        self.report_data['u_d'] = "%.4f" % self.data['num_u_d']
        self.report_data['u_lbd'] = "%.4f" % self.data['num_u_lbd']
        # 调用ReportWriter类
        RW = ReportWriter()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    ogr = Oscillograph()
