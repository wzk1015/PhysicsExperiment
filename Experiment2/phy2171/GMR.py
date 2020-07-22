import xlrd
from numpy import sqrt, abs, sin, cos, pi

import sys
sys.path.append('../..') # 如果最终要从main.py调用，则删掉这句
from GeneralMethod.PyCalcLib import Method, Simplified
from GeneralMethod.Report import Report


class Spectrometer:

    PREVIEW_FILENAME = "Preview.pdf"  # 预习报告模板文件的名称
    DATA_SHEET_FILENAME = "data.xlsx"  # 数据填写表格的名称
    REPORT_TEMPLATE_FILENAME = "GMR_empty.docx"  # 实验报告模板（未填数据）的名称
    REPORT_OUTPUT_FILENAME = "../../Report/Experiment2/2171Report.docx"  # 最后生成实验报告的相对路径

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
            print("该实验没有自动数据处理")
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            Method.start_file(self.REPORT_OUTPUT_FILENAME)
            print("Done!")

    '''
    填充实验报告
    调用Report类，将数据填入Word文档格式的实验报告中
    '''
    def fill_report(self):
        # 调用ReportWriter类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)


if __name__ == '__main__':
    mc = Spectrometer()
