# 密里根油滴实验（目前缺少不确定度计算的部分）
import GeneralMethod.GeneralMethod as gm
import math
import xlrd


class Millikan:
	def __init__(self):
		data = Millikan.get_data()
		[self.volt_list, self.time_list] = Millikan.get_exp_data()
		# volt_list是电压的一列，time_list是时间的8列表格
		self.ro1 = data['ro1']  # 油滴的密度
		self.ro2 = data['ro2']  # 空气的密度
		self.g = data['g']  # 重力加速度
		self.ita = data['ita']  # 空气粘度系数
		self.s = data['s']  # 油滴运动距离
		self.b = data['b']  # 修正系数
		self.p = data['p']  # 大气压强
		self.d = data['d']  # 极板间距
		self.e = data['e']  # 基本电荷量标准值
		Millikan.result(self)

	@staticmethod
	def get_data():
		# 从MillikanOilDrop.txt中获取需要的参数
		f = open("MillikanOilDrop/MillikanOilDrop.txt", "r")
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
		data = xlrd.open_workbook('MillikanOilDrop/MillikanOilDropData.xlsx')
		table = data.sheet_by_index(0)
		volt_list = table.col_values(colx=2, start_rowx=1, end_rowx=None)
		time_list = []
		# 第一列是电压（动态只有第一行平衡电压），后面8列是8组数据
		for i in range(18):
			time_list.append(table.row_values(rowx=i+1, start_colx=4, end_colx=10))
		return volt_list, time_list

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
			q_part3 = (1 / tf + 1 / tr) / u / pow(tf, 0.5)
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
			if i % 3 == 0:  # 处理静态法数据
				tf = gm.average(self.time_list[i])
				qc = Millikan.calculate_q(self, 0, tf, self.volt_list[i], 1)
				ec, en, err = Millikan.calculate_e(self, qc)
				res = [tf, qc, ec, en, err, 1]
				res_list.append(res)
			elif i % 3 == 1:    # 处理动态法数据
				tr = gm.average(self.time_list[i])
				tf = gm.average(self.time_list[i+1])
				qc = Millikan.calculate_q(self, tr, tf, 1.5 * self.volt_list[i], 2)
				ec, en, err = Millikan.calculate_e(self, qc)
				res = [tr, tf, qc, ec, en, err, 2]
				res_list.append(res)
		m = res_list
		for i in range(len(m)):
			if m[i][-1] == 1:
				print("第{}次密里根油滴实验静态法：".format(int((i+2)/2)))
				print("\t下降时间：\t{}".format(m[i][0]))
				print("\t油滴电荷量：\t{}\n\t基本电荷量：\t{}\n\t电荷数量：\t{}\n\t相对误差：\t{}\n".format(m[i][1], m[i][2], m[i][3], m[i][4]))
			else:
				print("第{}次密里根油滴实验动态法：".format(int((i+2)/2)))
				print("\t上升时间：\t{}\n\t下降时间：\t{}".format(m[i][0], m[i][1]))
				print("\t油滴电荷量：\t{}\n\t基本电荷量：\t{}\n\t电荷数量：\t{}\n\t相对误差：\t{}\n".format(m[i][2], m[i][3], m[i][4], m[i][5]))
		return

