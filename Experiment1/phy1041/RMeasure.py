import xlrd
import shutil
import os
from numpy import sqrt, abs

import sys
sys.path.append('../..')  # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Fitting, InstrumentError, Method
from reportwriter.ReportWriter import ReportWriter


class RMeasure:
    report_data_keys = [
        "y1", "y2", "y3", "y4", "y5", "y6", "y7", "y8",
        "x1", "x2", "x3", "x4", "x5", "x6", "x7", "x8",
        "y21", "y22", "y23", "y24", "y25", "y26", "y27", "y28",
        "x21", "x22", "x23", "x24", "x25", "x26", "x27", "x28",
        "xy1", "xy2", "xy3", "xy4", "xy5", "xy6", "xy7", "xy8",
        "ave_x", "ave_y", "b", "a", "r", "Rx", "ua_b", "ub_rx_rx", "ub_rx", "u_rx", "fin_Rx",
        "2r1", "2r2", "2u", "2r0", "2d", "2ri", "2fin_rg", "2ki",
        "fin_3rxh", "fin_u_3rxh_3rxh", "fin_u_3rxh", "fin_s_3_rxh",
        "3rs", "3v", "3d", "3Rg", "3ki", "u3_r", "u3_v", "u3_rs", "u3_d", "u3_rg", "u3_ki"
    ]
    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "data.xlsx"
    REPORT_TEMPLATE_FILENAME = "RMeasure_empty.docx"
    REPORT_OUTPUT_FILENAME = "RMeasure_out.docx"

    def __init__(self):
        self.data = {}  # 存放实验中的各个物理量
        self.uncertainty = {}  # 存放物理量的不确定度
        self.report_data = {}  # 存放需要填入实验报告的
        print("1041 电阻的测量\n1. 实验预习\n2. 数据处理")
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
            # './' is necessary when running this file, but should be removed if run main.py
            self.input_data("./"+self.DATA_SHEET_FILENAME)
            print("数据读入完毕，处理中......")
            # 计算物理量
            self.calc_data()
            # 计算不确定度
            # self.calc_uncertainty()
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            os.startfile(self.REPORT_OUTPUT_FILENAME)
            print("Done!")

    def input_data(self, filename):
        ws = xlrd.open_workbook(filename).sheet_by_name('RMeasure')
        # 从excel中读取数据
        list_x = []
        list_y = []
        list_2 = []
        list_3 = []
        list_r0 = []
        list_a0 = []
        list_r2 = []
        list_a2 = []
        list_r1 = []
        list_a1 = []
        list_rbox0 = []
        # list_y = ws.row_values(3,1)
        # list_x = ws.row_values(4,1)
        # list_2 = ws.row_values(11,1)
        # list_3 = ws.row_values(15,1)
        # list_a0 = ws.row_values(18,1)
        # list_r0 = ws.row_values(19,1)
        # list_a1 = ws.row_values(22,1)
        # list_r1 = ws.row_values(23,1)
        # list_a2 = ws.row_values(26,1)
        # list_r2 = ws.row_values(27,1)
        list_rbox0.append(float(ws.cell_value(20, 1)))
        list_rbox0.append(float(ws.cell_value(24, 1)))
        list_rbox0.append(float(ws.cell_value(28, 1)))

        num_len1 = int(ws.cell_value(17, 3))
        num_len2 = int(ws.cell_value(21, 3))
        num_len3 = int(ws.cell_value(25, 3))

        for col in range(1, ws.row_len(3)):
            list_y.append(float(ws.cell_value(3, col)))
            list_x.append(float(ws.cell_value(4, col)))
        for col in range(1, 6):
            list_2.append(float(ws.cell_value(11, col)))
            list_3.append(float(ws.cell_value(15, col)))
        # TODO 怎么读入不定长的表格数据？
        list_3.append(float(ws.cell_value(30, 2)))  # a1
        list_3.append(float(ws.cell_value(31, 2)))  # a2
        list_3.append(float(ws.cell_value(30, 1)))  # Nm1
        list_3.append(float(ws.cell_value(31, 1)))  # Nm2
        list_3.append(float(ws.cell_value(15, 6)))  # u3d
        list_3.append(float(ws.cell_value(15, 7)))  # u3rg
        # list_3.append(float(ws.cell_value(18, col)))#u3d
        list_3.append(float(ws.cell_value(15, 8)))  # u3ki

        for col in range(1, num_len1+1):
            list_a0.append(float(ws.cell_value(18, col)))
            list_r0.append(float(ws.cell_value(19, col)))
        for col in range(1, num_len2+1):
            list_a1.append(float(ws.cell_value(22, col)))
            list_r1.append(float(ws.cell_value(23, col)))
        for col in range(1, num_len3+1):
            list_a2.append(float(ws.cell_value(26, col)))
            list_r2.append(float(ws.cell_value(27, col)))

        self.data['list_a0'] = list_a0
        self.data['list_r0'] = list_r0
        self.data['list_a2'] = list_a2
        self.data['list_r2'] = list_r2
        self.data['list_a1'] = list_a1
        self.data['list_r1'] = list_r1
        self.data['list_rbox0'] = list_rbox0
        self.data['list_x'] = list_x
        self.data['list_y'] = list_y
        self.data['list_2'] = list_2
        self.data['list_3'] = list_3
        # print(list_x)

    def calc_data(self):
        # 实验一
        # TODO ua_rx和ub_rx的计算方式有点问题 （不知道那个k是个什么物理量啊）
        num_b, num_a, num_r = Fitting.linear(
            self.data['list_x'], self.data['list_y'], False)
        self.data['num_a'] = num_a
        self.data['num_b'] = num_b
        self.data['num_r'] = num_r
        num_rx = num_b
        list_k = []
        for i in range(0, 8):
            list_k.append(self.data['list_x'][i]/self.data['list_y'][i])
        # num_ua_b = Method.a_uncertainty(list_k)
        num_k = 8
        num_ave_x = Method.average(self.data['list_x'])
        num_ave_y = Method.average(self.data['list_y'])
        num_ua_b = num_b * (1 / (num_k-2) * ((1/num_r) ** 2 - 1)) ** (1/2)
        num_u_u = 0.00433
        num_u_i = 0.0000433
        num_ub_rx_rx = sqrt((num_u_u/num_ave_y) ** 2 +
                            (num_u_i/num_ave_x) ** 2)
        num_ub_rx = num_ub_rx_rx * num_rx
        num_u_rx = (num_ub_rx ** 2 + num_ua_b ** 2) ** (1/2)
        num_u_rx_1bit, pwr = Method.scientific_notation(num_u_rx)
        num_fin_rx = int(num_rx * (10 ** pwr)) / (10 ** pwr)
        self.data['num_ave_y'] = num_ave_y
        self.data['num_ave_x'] = num_ave_x
        self.data['num_ub_rx'] = num_ub_rx
        self.data['num_ub_rx_rx'] = num_ub_rx_rx
        self.data['num_u_u'] = num_u_u
        self.data['num_u_i'] = num_u_i
        self.data['num_rx'] = num_rx
        self.data['num_ua_b'] = num_ua_b
        self.data['num_u_rx'] = num_u_rx
        self.data['str_fin_rx'] = "%.0f±%.0f" % (num_fin_rx, num_u_rx_1bit)

        # 实验二
        num_2r1 = self.data['list_2'][0]
        num_2u = self.data['list_2'][1]
        num_2r0 = self.data['list_2'][2]
        num_2d = self.data['list_2'][3]
        num_2r2 = self.data['list_2'][4]
        num_2rg = num_2r2
        num_2ki = (num_2r1 * num_2u) / ((num_2r0 + num_2r1) * num_2rg * num_2d)
        num_dt_2r0 = InstrumentError.resistance_box(
            self.data['list_a0'], self.data['list_r0'], self.data['list_rbox0'][0])
        num_dt_2r1 = InstrumentError.resistance_box(
            self.data['list_a1'], self.data['list_r1'], self.data['list_rbox0'][1])
        num_dt_2r2 = InstrumentError.resistance_box(
            self.data['list_a2'], self.data['list_r2'], self.data['list_rbox0'][2])
        num_u_2r0 = num_dt_2r0 / sqrt(3)
        num_u_2r1 = num_dt_2r1 / sqrt(3)
        num_u_2r2 = num_dt_2r2 / sqrt(3)
        num_u_2u = 0.005*3/sqrt(3)
        num_u_2d = 1/sqrt(3)
        num_u_2r1_2r0 = sqrt((num_u_2r1) ** 2 + (num_u_2r0) ** 2)
        num_u_2ki_2ki = sqrt((num_u_2r1/num_2r1)**2+(num_u_2r1_2r0/(num_2r0+num_2r1))
                             ** 2+(num_u_2u/num_2u)**2+(num_u_2r2/num_2r2)**2+(num_u_2d/num_2d)**2)
        num_u_2ki = num_u_2ki_2ki * num_2ki
        num_u_2rg_1bit, pwr = Method.scientific_notation(num_u_2r2)
        num_fin_rg = int(num_rx * (10 ** pwr)) / (10 ** pwr)
        num_u_2ki_1bit, pwr = Method.scientific_notation(num_u_2ki)
        num_fin_ki = int(num_rx * (10 ** pwr)) / (10 ** pwr)
        self.data['str_fin_ki'] = "%.3f±%.3f" % (num_fin_ki, num_u_2ki_1bit)
        self.data['str_fin_rg'] = "%.0f±%.0f" % (num_fin_rg, num_u_2rg_1bit)
        self.data['num_dt_2r0'] = num_dt_2r0
        self.data['num_dt_2r1'] = num_dt_2r1
        self.data['num_dt_2r2'] = num_dt_2r2
        self.data['num_2r1'] = num_2r1
        self.data['num_2r2'] = num_2r2
        self.data['num_2r0'] = num_2r0
        self.data['num_2d'] = num_2d
        self.data['num_2u'] = num_2u
        self.data['num_u_2u'] = num_u_2u
        self.data['num_u_2d'] = num_u_2d
        self.data['num_u_2r1_2r0'] = num_u_2r1_2r0
        self.data['num_u_2ki_2ki'] = num_u_2ki_2ki
        self.data['num_u_2ki'] = num_u_2ki
        self.data['num_2ki'] = num_2ki

        # 实验三
        num_3rs = self.data['list_3'][0]
        num_3v = self.data['list_3'][1]
        num_3d = self.data['list_3'][2]
        num_3rg = self.data['list_3'][3]
        num_3ki = self.data['list_3'][4]
        num_u3_Nm1 = self.data['list_3'][7]
        num_u3_Nm2 = self.data['list_3'][8]
        num_u3_d = self.data['list_3'][9]
        num_u3_a1 = self.data['list_3'][5]
        num_u3_a2 = self.data['list_3'][6]
        num_u3_rg = self.data['list_3'][10]
        num_u3_ki = self.data['list_3'][11]

        num_3rxh = num_3rs/((num_3rg+num_3rs)*num_3ki)*num_3v/num_3d
        num_u3_rs = InstrumentError.electromagnetic_instrument(
            num_u3_a1, num_u3_Nm1) / sqrt(3)
        num_u3_v = InstrumentError.electromagnetic_instrument(
            num_u3_a2, num_u3_Nm2) / sqrt(3)
        num_u3_r = (num_u3_rg**2+num_u3_rs**2) ** (1/2)
        num_u3_rxh_rxh = ((num_u3_rs/num_3rs) ** 2+(num_u3_r/num_3rs) ** 2 +
                          (num_u3_v/num_3v)**2+(num_u3_ki/num_3ki)**2+(num_u3_d/num_3d)**2) ** (1/2)
        num_u3_rxh = num_u3_rxh_rxh * num_3rxh
        num_u3_rxh_1bit, pwr = Method.scientific_notation(num_u3_rxh)
        num_fin_rxh = int(num_3rxh * (10 ** pwr)) / (10 ** pwr)

        self.data['str_fin_rxh'] = "%.3f±%.3f" % (num_fin_rxh, num_u3_rxh_1bit)
        self.data['num_3rxh'] = num_3rxh
        self.data['num_u3_rs'] = num_u3_rs
        self.data['num_u3_v'] = num_u3_v
        self.data['num_u3_r'] = num_u3_r
        self.data['num_u3_rxh_rxh'] = num_u3_rxh_rxh
        self.data['num_u3_rxh'] = num_u3_rxh
        self.data['num_3rs'] = num_3rs
        self.data['num_3v'] = num_3v
        self.data['num_3d'] = num_3d
        self.data['num_3rg'] = num_3rg
        self.data['num_3ki'] = num_3ki
        self.data['num_u3_Nm1'] = num_u3_Nm1
        self.data['num_u3_Nm2'] = num_u3_Nm2
        self.data['num_2ki'] = num_2ki
        self.data['num_u3_d'] = num_u3_d
        self.data['num_u3_a1'] = num_u3_a1
        self.data['num_u3_a2'] = num_u3_a2
        self.data['num_u3_rg'] = num_u3_rg
        self.data['num_u3_ki'] = num_u3_ki

    def fill_report(self):
        for i, x_i in enumerate(self.data['list_x']):
            self.report_data["x"+str(i + 1)] = "%.5f" % (x_i)
            self.report_data["x2"+str(i + 1)] = "%.5f" % (x_i ** 2)
        for i, y_i in enumerate(self.data['list_y']):
            self.report_data["y"+str(i + 1)] = "%.5f" % (y_i)
            self.report_data["y2"+str(i + 1)] = "%.5f" % (y_i ** 2)
        for i in range(0, 8):
            self.report_data["xy"+str(i+1)] = (self.data['list_x']
                                               [i])*(self.data['list_y'][i])
        self.report_data['ave_y'] = "%.4f" % self.data['num_ave_y']
        self.report_data['ave_x'] = "%.4f" % self.data['num_ave_x']
        self.report_data['b'] = "%.4f" % self.data['num_b']
        self.report_data['a'] = "%.4f" % self.data['num_a']
        self.report_data['r'] = "%.4f" % self.data['num_r']
        self.report_data['Rx'] = "%.4f" % self.data['num_rx']
        self.report_data['ua_b'] = "%.4f" % self.data['num_ua_b']
        self.report_data['ub_rx_rx'] = "%.4f" % self.data['num_ub_rx_rx']
        self.report_data['ub_rx'] = "%.4f" % self.data['num_ub_rx']
        self.report_data['u_rx'] = "%.4f" % self.data['num_u_rx']
        self.report_data['fin_Rx'] = self.data['str_fin_rx']
        self.report_data['dt_2r0'] = self.data['num_dt_2r0']
        self.report_data['dt_2r1'] = self.data['num_dt_2r1']
        self.report_data['dt_2r2'] = self.data['num_dt_2r2']
        self.report_data['2u'] = self.data['num_u_2u']
        self.report_data['2d'] = self.data['num_u_2d']
        self.report_data['2r2'] = self.data['num_2r2']
        self.report_data['2r1'] = self.data['num_2r1']
        self.report_data['2r0'] = self.data['num_2r0']

        self.report_data['u_2r1_2r0'] = self.data['num_u_2r1_2r0']
        self.report_data['u_2ki_2ki'] = self.data['num_u_2ki_2ki']
        self.report_data['u_2ki'] = self.data['num_u_2ki']
        self.report_data['2ki'] = self.data['num_2ki']
        self.report_data['fin_ki'] = self.data['str_fin_ki']
        self.report_data['2fin_rg'] = self.data['str_fin_rg']
        self.report_data['2ub_r2'] = self.data['num_dt_2r2']
        self.report_data['2fin_rg'] = self.data['num_dt_2r2']

        self.report_data['3rs'] = self.data['num_3rs']
        self.report_data['3v'] = self.data['num_3v']
        self.report_data['3d'] = self.data['num_3d']
        self.report_data['3rg'] = self.data['num_3rg']
        self.report_data['3ki'] = self.data['num_3ki']
        # self.report_data['u3_Nm1'] = self.data['num_u3_Nm1']
        # self.report_data['u3_Nm2'] = self.data['num_u3_Nm2']
        # self.report_data['fin_rg'] = self.data['num_u3_d']
        # self.report_data['fin_rg'] = self.data['num_u3_a1']
        # self.report_data['fin_rg'] = self.data['num_u3_a2']
        self.report_data['u3_rg'] = self.data['num_u3_rg']
        self.report_data['u3_ki'] = self.data['num_u3_ki']
        self.report_data['fin_3rxh'] = self.data['str_fin_rxh']
        self.report_data['3rxh'] = self.data['num_3rxh']
        self.report_data['u3_rs'] = self.data['num_u3_rs']
        self.report_data['u3_v'] = self.data['num_u3_v']
        self.report_data['u3_r'] = self.data['num_u3_r']
        self.report_data['u3_d'] = self.data['num_u3_d']
        self.report_data['fin_u_3rxh_3rxh'] = self.data['num_u3_rxh_rxh']
        self.report_data['fin_u_3rxh'] = self.data['num_u3_rxh']
        self.report_data['fin_s_3_rxh'] = self.data['str_fin_rxh']
        # self.report_data['u3_rxh'] = self.data['num_u3_rxh']

        RW = ReportWriter()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME,
                       self.REPORT_OUTPUT_FILENAME)


if __name__ == '__main__':
    rms = RMeasure()
