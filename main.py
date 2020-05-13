# 主程序，只包括选择实验和运行函数，以后会加上说明一类的
import MillikanOilDrop.MillikanOilDrop as Millikan
import SolarBattery.SolarBattery as sb
import Frank_Hertz.frank_hertz as fh


if __name__ == '__main__':
	# 主程序，只引用模块
	try:
		print("目前可计算的实验：")
		print("\t1、密里根油滴实验")
		print("\t2、太阳能电池特性实验")
		print("\t3、弗兰克赫兹实验")
		exp = input("请输入实验序号：")
		if exp == "1":
			Millikan.Millikan()
		elif exp == "2":
			sb.SolarBattery()
		elif exp == "3":
			fh.FrankHertz()
		else:
			print("很抱歉。暂时没有相应的数据处理程序。")
	except ValueError:
		print("请输入一个数字")
	input("点击任意键退出")
