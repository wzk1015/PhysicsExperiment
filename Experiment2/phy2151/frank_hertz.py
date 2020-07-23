#!/usr/bin/env python
# -*- coding: utf-8 -*-
# author: seven_lit 
# name:frank_hertz.py time: 2020/5/13

import xlrd
import numpy
import math
import os
import  sys
 #若最终要从main.py 调用，则删掉这句
from GeneralMethod.PyCalcLib import Method
from reportwriter.ReportWriter import ReportWriter

from win32com.client.gencache import EnsureDispatch
import sys
xl = EnsureDispatch("Word.Application")
print(sys.modules[xl.__module__].__file__)
class FrankHertz:
	# 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#
	report_data_keys = [
		"1","2","3","4","5","6",#每次波峰的值
		"3d-1","3d-2","3d-3",#逐差法：3Δd(本程序按波峰逐差，如有需要也可选择波谷，注意输入输出的结果)
		"Ua","Ub","U" ,"Ave"#a类不确定度， b类不确定度， 合成不确定度， 波峰（谷）平均数
		"final" #最终结果
		"relative_error"  #相对误差
	]
	PREVIEW_FILENAME = "Preview.pdf"
	DATA_SHEET_FILENAME = "data_frank_hertz.xlsx"
	REPORT_TEMPLATE_FILENAME = "FrankHertz_empty.docx"
	REPORT_OUTPUT_FILENAME = "FrankHertz_out.docx"

	def __init__(self):
		self.data = {}  # 存放实验中的各个物理量
		self.uncertainty = {}  # 存放物理量的不确定度
		self.report_data = {}  # 存放需要填入实验报告的
		print("2151 弗兰克赫兹实验\n1. 实验预习\n2. 数据处理")
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
			self.input_data(
				"./" + self.DATA_SHEET_FILENAME)  # './' is necessary when running this file, but should be removed if run main.py
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
		#  文件名
		'''
		从excel表格中读取数据
		@param filename: 输入excel的文件名
		@return none
		'''
	def input_data(self,filename):
		ws = xlrd.open_workbook(filename).sheet_by_name('FrankHertz')
		#从excel中读取数据
		list_d = []
		for col in range(2,8):
			list_d.append(float(ws.cell_value(1,col)))
		self.data['list_d'] = list_d



	def calc_data(self):
		list_dif_d, ave_d = Method.successive_diff(self.data['list_d'])
		self.data['list_dif_d'] = (1/3)*list_dif_d
		self.data['ave_d'] = (1/3)*ave_d



	def calc_uncertainty(self):
		Ua =Method.a_uncertainty(self.data['list_dif_d']) # a类不确定度
		Ub = 0.1 / (numpy.sqrt(3))  # b类不确定度
		U = numpy.sqrt(numpy.square(Ua) + numpy.square(Ub))  # 总的不确定度
		self.data.update({"Ua":Ua, "Ub":Ub, "U":U})
		ave_d = self.data['ave_d']
		#res_final, unc_final, pwr = Method.final_expression(ave_d,U)
		#self.data['final'] = "(%s±%s)*10e(%s)"%(res_final,unc_final,pwr)

		#base_U, pwr = Method.scientific_notation(U) #将不确定度转化为只有一位有效数字的科学计数法
		#volt_final = int(ave_d *(10**pwr))/(10 ** pwr) #对物理量保留有效数字，截断处理
		self.data['final'] = "%.1f±%.1f" % (ave_d,U)
		self.data['relative_error'] = abs((float(int(ave_d*10)/10-11.55))/11.55)

	'''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''
	def test(self):
		ave_d = self.data['ave_d']
		print(ave_d)

	def fill_report(self):
		# 表格：原始数据d
		for i, d_i in enumerate(self.data['list_d']):
			self.report_data[str(i + 1)] = "%.1f" % (d_i)  # 一定都是字符串类型
		# 表格：逐差法计算3Δd
		for i, dif_d_i in enumerate(self.data['list_dif_d']):
			self.report_data["3d-%d" % (i + 1)] = "%.1f" % (dif_d_i)
		# 最终结果
		self.report_data['final'] = self.data['final']
		# 将各个变量以及不确定度的结果导入实验报告，在实际编写中需根据实验报告的具体要求设定保留几位小数
		self.report_data['Ave'] = "%.1f" % self.data['ave_d']
		self.report_data['Ua'] = "%.2f" % self.data['Ua']
		self.report_data['Ub'] = "%.3f" % self.data['Ub']
		self.report_data['U'] = "%.1f" % self.data['U']
		self.report_data['relative_error'] = "%.3f" % self.data['relative_error']

		# 调用ReportWriter类
		RW = ReportWriter()
		RW.load_replace_kw(self.report_data)
		RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)


if __name__ == '__main__':
	mc = FrankHertz()

