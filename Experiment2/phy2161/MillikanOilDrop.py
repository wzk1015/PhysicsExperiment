# 密里根油滴实验（目前缺少不确定度计算的部分）
import GeneralMethod.PyCalcLib as gm
import math
import xlrd
import shutil
import os
from numpy import sqrt, abs
import sys
sys.path.append('../..')   # 如果最终要从main.py 调用，则删掉这句
from reportwriter.ReportWriter import ReportWriter


class Millikan:
	# 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
	report_data_keys = [
		"V1", "V2", "V3", "V4", "V5", "V6", #六组数据各自对应的电压
		"1-1","1-2","1-3","1-4","1-5",#每组测5次
		"2-1","2-2","2-3","2-4","2-5",
		"3-1","3-2","3-3","3-4","3-5",
		"4-1","4-2","4-3","4-4","4-5"
		"5-1","5-2","5-3","5-4","5-5"
		"6-1","6-2","6-3","6-4","6-5"
		"Ave1", "Ave2", "Ave3", "Ave4", "Ave5",  # 静态法平均时间
		"q1", "q2", "q3","q4","q5","q6",  # 六组计算电荷量
		"n1","n2","n3","n4","n5","n6", #每组对应的单位电荷个数
		"e1","e2","e3","e4","e5","e6",  # 100圈光程差d的不确定度
		"rel_err1","rel_err2","rel_err3","rel_err4","rel_err5","rel_err6" #每组相对误差
		"rel_err"# 总相对误差
		"final"  # 最终结果
	]

	PREVIEW_FILENAME = "Preview.pdf"
	DATA_SHEET_FILENAME = "Data.xlsx"
	REPORT_TEMPLATE_FILENAME = "MillikanOilDrop_empty.docx"
	REPORT_OUTPUT_FILENAME = "MillikanOilDrop_out.docx"


	def __init__(self):
		self.data = {} #存放实验的各个物理量
		self.uncertainty = {}  # 存放物理量的不确定度
		self.report_data = {}  # 存放需要填入实验报告的
		# volt_list是电压的一列，time_list是时间的8列表格
		print("2161 密立根油滴实验\n1. 实验预习\n2. 数据处理")
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
			data = Millikan.get_data()
			[self.volt_list, self.time_list,self.time_Ave] = Millikan.get_exp_data()

			self.ro1 = data['ro1']  # 油滴的密度
			self.ro2 = data['ro2']  # 空气的密度
			self.g = data['g']  # 重力加速度
			self.ita = data['ita']  # 空气粘度系数
			self.s = data['s']  # 油滴运动距离
			self.b = data['b']  # 修正系数
			self.p = data['p']  # 大气压强
			self.d = data['d']  # 极板间距
			self.e = data['e']  # 基本电荷量标准值
			print("数据读入完毕，处理中......")
			# 计算电量
			Millikan.result(self)
			print("正在生成实验报告......")
			#生成实验报告
			self.fill_report()
			print("实验报告生成完毕，正在打开......")
			os.startfile(self.REPORT_OUTPUT_FILENAME)
			print("Done!")



	@staticmethod
	def get_data():
		# 从MillikanOilDrop.txt中获取需要的参数
		f = open("MillikanOilDrop.txt", "r")
		data = f.readlines()
		data_table = []
		data_dict = {}
		for i in range(len(data)):
			data_table.append(data[i].split(" "))
			# 这里是以空格分割，因此数据和符号之间一定要有空格，这里考虑到用途，不会放到excel表格里
		data_dict['ro1'] = float(data_table[0][2])
		data_dict['ro2'] = float(data_table[1][2])
		data_dict['g'] = float(data_table[2][2])
		data_dict['ita'] = float(data_table[3][2])
		data_dict['s'] = float(data_table[4][2])
		data_dict['b'] = float(data_table[5][2])
		data_dict['p'] = float(data_table[6][2])
		data_dict['d'] = float(data_table[7][2])
		data_dict['e'] = float(data_table[8][2])
		f.close()
		return data_dict

	@staticmethod
	def get_exp_data():
		# 从MillikanOilDropData.xlsx中获取实验数据
		data = xlrd.open_workbook('Data.xlsx')
		table = data.sheet_by_index(0)
		volt_list = table.col_values(colx=2, start_rowx=1, end_rowx=7)

		time_list = []
		time_Ave = []
		# 第一列是电压（动态只有第一行平衡电压），后面8列是8组数据
		for i in range(6):
			time_list.append(table.row_values(rowx=i+1, start_colx=3, end_colx=8))

		for i in range(6):
			sum = 0
			for j in range(5):
				sum = time_list[i][j]+sum
			time_Ave.append(sum/5.)

		return volt_list, time_list,time_Ave

	def calculate_r0(self, vf):
		# 计算修正需要的r0
		r0 = pow(9 * self.ita * vf / ((self.ro1 - self.ro2) * self.g * 2), 0.5)
		return r0

	def calculate_q(self, tr, tf, u, model):
		# 计算q（油滴带的电荷量）
		pi = math.pi
		vf = self.s/tf
		r0 = Millikan.calculate_r0(self, vf)
		# 把q分成四部分计算，否则就太长了。另外动态和静态要分开算。
		q_prat1 = 9 * pow(2, 0.5) * pi * self.d
		q_part2 = pow(self.ita * self.s, 1.5) / pow((self.ro1 - self.ro2) * self.g, 0.5)
		if model == 1:
			q_part3 = 1 / u / pow(tf, 1.5)
		else:
			q_part3 = (1 / tf + 1 / tr) / u / pow(tf, 0.5)          #动态法残余内容，如果有需要可以修改为带有动态法的版本
		q_part4 = pow(1 / (1 + self.b/(self.p*r0)), 1.5)
		q = q_prat1 * q_part2 * q_part3 * q_part4
		return q

	def calculate_e(self, q):
		# 计算最后结果e（基本电荷量）以及相对误差。
		e = self.e
		n = round(q/e)
		e1 = q/n
		error = (e1-e) / e
		return e1, n, error

	def result(self):
		# 将所有部分整合。求得最终结果并展示
		res_list = []
		for i in range(len(self.volt_list)):
			tf = self.time_Ave[i]
			qc = Millikan.calculate_q(self, 0, tf, self.volt_list[i], 1)
			ec, en, err = Millikan.calculate_e(self, qc)
			res = [tf, qc, ec, en, err, 1]
			res_list.append(res)
		m = res_list
		self.data['m']= m
		self.data['final'] = (m[0][2]+m[1][2]+m[2][2]+m[3][2]+m[4][2]+m[5][2])/6.
		self.data['rel_err'] = abs((self.data['final']-self.e)/self.e)
		return

	def fill_report(self):
		# 表格：原始数据d
		for i, volt in enumerate(self.volt_list):
			self.report_data['V%d'%(i+1)] = "%d" % (volt) # 静态法电压
		for i in range(6):
			for j in range(5):
				self.report_data['%d-%d'%(i+1,j+1)] = "%.2f"%(self.time_list[i][j])  #时间
		m = self.data['m']
		for i in range(len(m)):

			self.report_data['q%d'%(i+1)] = "%.2f*10e(-19)"%(m[i][1]*(10**19))  #油滴电荷量
			self.report_data['e%d'%(i+1)] = "%.2f*10e(-19)"%(m[i][2]*(10**19))  #实验e数值
			self.report_data['n%d'%(i+1)] = "%d"%(m[i][3])  #油滴单位电荷数目
			self.report_data['rel_err%d'%(i+1)] = "%.3f"%(abs(m[i][4]))  #相对误差

		for i,Ave in enumerate(self.time_Ave):
			self.report_data['Ave%d'%(i+1)] = "%.2f"%(Ave)  #时间均值
		self.report_data['final'] = "%.2f*10e(-19)"%((self.data['final'])*(10**19)) #最后结果不带不确定度
		self.report_data['rel_err'] = "%.3f"%(self.data['rel_err'])

		# 调用ReportWriter类
		RW = ReportWriter()
		RW.load_replace_kw(self.report_data)
		RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    mc = Millikan()
