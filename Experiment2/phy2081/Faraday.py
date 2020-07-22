import xlrd
# from xlutils.copy import copy as xlscopy
import shutil
import os
import math
import sys
o_path = os.path.abspath(os.path.join(os.getcwd(), "../..")) # 调用库需要返回到当前目录的上上级
sys.path.append(o_path) # 如果最终要从main.py调用，则删掉这句

from GeneralMethod.PyCalcLib import Method, Fitting
from GeneralMethod.Report import Report


class Faraday:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
    report_data_keys = [
        "1", "2", "3", "4", "5", "6", "7", "8", "9", "10",
        "11", "12", "13", "14", "15", "16", "17", "18", "19", "20",
        "21", "22", "23", "24", "25", "26", "27", "28", "29", "30",
        "31", # 记录的磁感应强度
        "a", "b", "r", # 线性拟合直线信息
        "ua_a", "ua_b", #线性拟合参数不确定度
        "t1_1",	"t1_2",	"t1_3",	"t1_4",
        "t2_1",	"t2_2",	"t2_3",	"t2_4", # 记录的角度
        "D", # 记录的距离
        "t_1", "t_2", "t_3", "t_4", # 计算得到的偏转角
        "B_1", "B_2", "B_3", "B_4", # 计算得到的磁感应强度
        "V_1", "V_2", "V_3", "V_4", "V_avg", # 计算得到的费尔德常数
        "ua_B", "ub_theta", "u_V1", "u_V2", "u_V3", "u_V4", "u_Vavg", # 不确定度 
    ]
    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "data.xlsx"
    REPORT_TEMPLATE_FILENAME = "Faraday_empty.docx"
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment2/2081Report.docx"

    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {} # 存放需要填入实验报告的
        print("2081 法拉第磁光效应实验\n1. 实验预习\n2. 数据处理")
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

    '''
    从excel表格中读取数据
    @param filename: 输入excel的文件名
    @return none
    '''
    def input_data(self, filename):
        # 从excel第一个工作簿中读取数据
        ws = xlrd.open_workbook(filename).sheet_by_name('Faraday')       
        Current1 = []
        Magnetic_induction1 = []
        for col in range(1, 32):
            Current1.append(float(ws.cell_value(0, col))) # 从excel取出来的数据，加个类型转换靠谱一点
            Magnetic_induction1.append(float(ws.cell_value(1, col)))
        X = [] # 取电流0~1.2A的范围做线性拟合，X是电流，Y是磁感应强度
        Y = []
        for i in range(0, 16):
            X[i] = Current1[i]
            Y[i] = Magnetic_induction1[i]
        self.data['Current1'] = Current1 # 存储从表格中读入的数据
        self.data['Magnetic_induction1'] = Magnetic_induction1
        self.data['X'] = X
        self.data['Y'] = Y
        # 从excel第二个工作簿中读取数据
        ws2 = xlrd.open_workbook(filename).sheet_by_name('Faraday2')
        Current2 = []
        Theta1 = []
        Theta2 = []       
        Distance = float(ws2.cell_value(7, 1))
        for col in range(1, 5):
            Current2.append(float(ws2.cell_value(0, col)))
            Theta1.append(float(ws2.cell_value(1, col)))
            Theta2.append(float(ws2.cell_value(2, col)))
        self.data['Current2'] = Current2 # 存储从表格中读入的数据        
        self.data['Theta1'] = Theta1
        self.data['Theta2'] = Theta2
        self.data['D'] = Distance
    
    '''
    进行数据处理
    由于2081实验的数据处理非常非常简单，为节约代码量，将全部数据处理放在一个函数内完成.
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    def calc_data(self):
        # 计算拟合直线的a, b(y = ax + b)和相关系数r
        [a, b, r] = Fitting.linear(self.data['X'], self.data['Y'])
        self.data['a'] = a
        self.data['b'] = b
        self.data['r'] = r
        # 计算磁光效应偏转角度，利用上面计算出的a, b拟合出磁感应强度，最终算出费尔德常数
        Theta = []
        n = len(self.data['Theta1'])
        for i in range(0, n):
            Theta.append(self.data['Theta2'][i] - self.data['Theta1'][i])
        self.data['Theta'] = Theta
        Magnetic_induction2 = []
        for i in range(0, n):
            Magnetic_induction2.append(a * self.data['Current2'][i] + b)
        self.data['Magnetic_induction2'] = Magnetic_induction2
        V = []
        for i in range(0, n):
            V.append((Theta[i] * math.pi / 180) / (Magnetic_induction2[i] * self.data['D']) * 1e6)
        self.data['V'] = V
        
    '''
    计算所有的不确定度
    '''
    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    def calc_uncertainty(self):
        # 计算线性拟合的A类不确定度
        Sigma = 0
        for i in range(0, len(self.data['X'])):
            Sigma += (self.data['Y'][i] - (self.data['b'] + self.data['a'] * self.data['X'][i]))**2
        k = len(self.data['X'])
        ua_B = (Sigma / (k - 2))**(1 / 2)
        ua_a = self.data['b'] * (((1 / self.data['r']**2 - 1) / (k - 2))**(1 / 2))
        X_square = []
        for i in range(0, k):
            X_square[i] = self.data['X'][i]**2
        ua_b = ua_a * (Method.average(X_square)**(1 / 2))
        self.data.update({"ua_B":ua_B, "ua_a":ua_a, "ua_b":ua_b})

        # 求theta的B类不确定度，分别合成不同V的不确定度，计算加权平均后V_avg的不确定度
        ub_theta = 1 / math.sqrt(3) * math.pi / 180
        u_V = []
        for i in range(0, 4):
            u_V.append(self.data['V'][i] * math.sqrt(1 / self.data['Theta'][i]**2 * ub_theta**2 + 1 / self.data['Magnetic_induction2'][i]**2 * ua_B**2))
        u_V1 = u_V[0]
        u_V2 = u_V[1]
        u_V3 = u_V[2]
        u_V4 = u_V[3]
        u2_Vavg = 1 / (u_V1**(-2) + u_V2**(-2) + u_V3**(-2) + u_V4**(-2))
        u_Vavg = math.sqrt(u2_Vavg)
        V_avg = u2_Vavg * (self.data['V'][0] / u_V1**2 + self.data['V'][1] / u_V2**2 + self.data['V'][2] / u_V3**2 + self.data['V'][3] / u_V4**2)
        self.data.update({"ub_theta":ub_theta, "u_V1":u_V1, "u_V2":u_V2, "u_V3":u_V3, "u_V4":u_V4, "u_Vavg":u_Vavg, "V_avg":V_avg})

    '''
    填充实验报告
    调用Report类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 表格：1原始数据
        for i, b_i in enumerate(self.data['Magnetic_induction1']):
            self.report_data[str(i + 1)] = "%d" % (b_i) # 一定都是字符串类型
        # 1数据处理
        self.report_data['a'] = "%.5f" % self.data['a']
        self.report_data['b'] = "%.5f" % self.data['b']
        self.report_data['r'] = "%.8f" % self.data['r']
        # 1不确定度
        self.report_data['ua_B'] = "%.5f" % self.data['ua_B']
        self.report_data['ua_a'] = "%.5f" % self.data['ua_a']
        self.report_data['ua_b'] = "%.5f" % self.data['ua_b']
        # 表格：2原始数据
        for i, t1_i in enumerate(self.data['Theta1']):
            self.report_data["t1_%d" % (i + 1)] = "%d" % (t1_i)
        for i, t2_i in enumerate(self.data['Theta2']):
            self.report_data["t2_%d" % (i + 1)] = "%d" % (t2_i)
        self.report_data['D'] = "%.2f" % self.data['D']
        # 2数据处理
        for i, t_i in enumerate(self.data['Theta']):
            self.report_data["t_%d" % (i + 1)] = "%d" % (t_i)
        for i, B_i in enumerate(self.data['Magnetic_induction2']):
            self.report_data["B_%d" % (i + 1)] = "%.2f" % (B_i)
        for i, V_i in enumerate(self.data['V']):
            self.report_data["V_%d" % (i + 1)] = "%.2f" % (V_i)
        # 2不确定度
        self.report_data['ub_theta'] = "%.5f" % self.data['ub_theta']
        self.report_data['u_V1'] = "%.3f" % self.data['u_V1']
        self.report_data['u_V2'] = "%.3f" % self.data['u_V2']
        self.report_data['u_V3'] = "%.3f" % self.data['u_V3']
        self.report_data['u_V4'] = "%.3f" % self.data['u_V4']
        self.report_data['u_Vavg'] = "%.4f" % self.data['u_Vavg']
        self.report_data['V_avg'] = "%.3f" % self.data['V_avg']
        # 调用Report类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    fa = Faraday()
