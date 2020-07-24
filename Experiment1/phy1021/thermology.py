import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
from numpy import sqrt, abs

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Fitting
from GeneralMethod.PyCalcLib import Method
from reportwriter.ReportWriter import ReportWriter

class thermology:
    report_data_keys = [
        '1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20',
        '21','22','23','24','25','26','27','28','29',
        '101','102','103','104','105','106','107','108','109','110','111','112','113','114','115','116','117',
        '118','119','120','121','122','123','124','125','126','127','128','129',
        'L','K','J',
        'Ua','UJ'
    ]

    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "data.xlsx"
    REPORT_TEMPLATE_FILENAME = "thermology_empty.docx"
    REPORT_OUTPUT_FILENAME = "thermology_out.docx"
    
    
    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {} # 存放需要填入实验报告的
        print("1021 测量水的溶解热+焦耳热功当量\n1. 实验预习\n2. 数据处理")
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
            self.calc_data1()
            self.calc_data2()
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
        ws = xlrd.open_workbook(filename).sheet_by_name('thermology1')
        # 从excel中读取数据
        list_time = []
        list_resistance = []
        list_temperature = []
        list_weight = []
        for row in [1, 4, 7]:
            for col in range(1, 8):
                list_time.append(float(ws.cell_value(row, col)))  #时间
        self.data['list_time'] = list_time
        for row in [2, 5, 8]:
            for col in range(1, 8):
                list_resistance.append(float(ws.cell_value(row, col)))  #电阻值
        self.data['list_resistance'] = list_resistance
        for row in [3, 6, 9]:
            for col in range(1, 8):
                list_temperature.append(float(ws.cell_value(row, col)))   #温度
        self.data['list_temperature'] = list_temperature
        
        col = 1
        for row in range(10, 14):
            list_weight.append(float(ws.cell_value(row,col)))    #几种质量
        self.data['list_weight'] = list_weight
        row = 14
        temp_ice = float(ws.cell_value(row, col))    #冰温度
        self.data['temp_ice'] = temp_ice
        row = 15
        temp_env = float(ws.cell_value(row, col))     #环境温度
        
        self.data['temp_env'] = temp_env

        ws = xlrd.open_workbook(filename).sheet_by_name('thermology2')
        list_time2 = []
        list_resistance2 = []
        list_temperature2 = []
        for row in [1, 4, 7, 10]:
            for col in range(1, 9):
                if isinstance(ws.cell_value(row, col), str):
                    continue
                else:
                    list_time2.append(float(ws.cell_value(row, col)))
        self.data['list_time2'] = list_time2
        for row in [2, 5, 8, 11]:
            for col in range(1, 9):
                if isinstance(ws.cell_value(row, col), str):
                    continue
                else:
                    list_resistance2.append(float(ws.cell_value(row, col)))
        self.data['list_resistance2'] = list_resistance2
        for row in [3, 6, 9, 12]:
            for col in range(1, 9):
                if isinstance(ws.cell_value(row, col), str):
                    continue
                else:
                    list_temperature2.append(float(ws.cell_value(row, col)))
        self.data['list_temperature2'] = list_temperature2
        
        
        col = 1
        row = 13
        temp_env2 = float(ws.cell_value(row, col))
        self.data['temp_env2'] = temp_env2
        row = 14
        voltage_begin = float(ws.cell_value(row, col))
        self.data['voltage_begin'] = voltage_begin
        row = 15
        voltage_end = float(ws.cell_value(row, col))
        self.data['voltage_end'] = voltage_end
        row = 16
        resitence = float(ws.cell_value(row, col))
        self.data['resitence'] = resitence

        self.data['c1'] = 0.389e3
        self.data['c2'] = 0.389e3
        self.data['c0'] = 4.18e3
        self.data['ci'] = 1.80e3

    def calc_data1(self):
        list_weight = self.data['list_weight']
        list_time1 = self.data['list_time']
        list_temperature1 = self.data['list_temperature']
        temp_ice = self.data['temp_ice']
        temp_env = self.data['temp_env']
        c1 = self.data['c1']
        c2 = self.data['c2']
        c0 = self.data['c0']
        ci = self.data['ci']
        
        m_water = list_weight[1] - list_weight[0]
        m_ice = list_weight[2] - list_weight[1]
        
        list_graph = Fitting.linear(list_time1, list_temperature1, show_plot=False)
        self.data['list_graph'] = list_graph
        
        

        temp_begin = list_graph[0] * list_time1[0] + list_graph[1]    
        temp_end = list_graph[0] * list_time1[(len(list_time1)-1)] + list_graph[1]
        self.data['temp_begin'] = temp_begin
        self.data['temp_end'] = temp_end
        self.data['m_water'] = m_water
        self.data['m_ice'] = m_ice
        
        '''
        print(temp_begin)
        print('\n',temp_end)
        print('\n',m_water)
        print('\n',m_ice)
        print('!1!\n',c0*m_water*0.001+c1*list_weight[3]*0.001+c2*(list_weight[0]-list_weight[3])*0.001)
        print('\n!2!\n',temp_begin-temp_end)
        print('\n!3!\n',c0*temp_end)
        print('\n!4!\n',ci*temp_ice)
        '''
        
        L = 1/(m_ice*0.001) * (c0*m_water*0.001+c1*list_weight[3]*0.001+c2*(list_weight[0]-list_weight[3])*0.001) * (temp_begin-temp_end)- c0*temp_end + ci*temp_ice
        K = c0 * m_water*0.001 * (list_temperature1[15]-list_temperature1[8]) / ((list_time1[15]-list_time1[8])*(list_temperature1[15]-temp_env))
        self.data['L'] = L
        self.data['K'] = K

    def calc_data2(self):
        c1 = self.data['c1']
        c0 = self.data['c0']
        list_temperature2 = self.data['list_temperature2']
        list_weight = self.data['list_weight']
        temp_env2 = self.data['temp_env2']
        list_time2 = self.data['list_time2']
        voltage_begin = self.data['voltage_begin']
        voltage_end = self.data['voltage_end']
        resitence = self.data['resitence'] 
        
        
        
        m_water = list_weight[1] - list_weight[0]
        list_x = []
        list_y = []

        for i in range(len(list_temperature2)):
            if i==len(list_temperature2)-1:
                break
            list_x.append((list_temperature2[i]+list_temperature2[i+1])/2-temp_env2)
        for i in range(len(list_temperature2)):
            if i == len(list_temperature2)-1:
                break
            list_y.append((list_temperature2[i+1]-list_temperature2[i])/((list_time2[i+1]-list_time2[i])*60))
        self.data['list_x'] = list_x
        self.data['list_y'] = list_y
        list_graph2 = Fitting.linear(list_x, list_y, show_plot=False)
        self.data['list_graph2'] = list_graph2
        J = ((voltage_begin+voltage_end)/2)**2/(list_graph2[1]*resitence*(c0*m_water*0.001+c1*list_weight[3]*0.001+64.38))
        self.data['J'] = J
        
        '''
        print('b',list_graph2[0])
        print('\n a',list_graph2[1])
        print('\n r',list_graph2[2])  
        '''

    def calc_uncertainty(self):
        list_a = []
        list_x = self.data['list_x']
        list_y = self.data['list_y']
        list_graph2 = self.data['list_graph2']
        voltage_begin = self.data['voltage_begin']
        voltage_end = self.data['voltage_end']
        resitence = self.data['resitence'] 
        c1 = self.data['c1']
        c0 = self.data['c0']
        list_weight = self.data['list_weight']
        m_water = list_weight[1] - list_weight[0]
        for i in range(len(list_x)):
            list_a.append(list_y[i]-list_graph2[1]*list_x[i])
        self.data['list_a'] = list_a
        Ua = Method.a_uncertainty(self.data['list_a'])
        self.data['Ua'] = Ua
        UJ = abs(((voltage_begin+voltage_end)/2)**2/(Ua*resitence*(c0*m_water*0.001+c1*list_weight[3]*0.001 + 64.38)))
        self.data['UJ'] = UJ

    def fill_report(self):
        # 表格：xy
        for i, x_i in enumerate(self.data['list_x']):
            self.report_data[str(i + 1)] = "%.5f" % (x_i)

        for i, y_i in enumerate(self.data['list_y']):
            self.report_data[str(i + 101)] = "%.5f" % (y_i)
            # 最终结果
        self.report_data['L'] = self.data['L']
        self.report_data['K'] = self.data['K']
        self.report_data['J'] = self.data['J']
        self.report_data['Ua'] = self.data['Ua']
        self.report_data['UJ'] = self.data['UJ']

        RW = ReportWriter()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    mc = thermology()






