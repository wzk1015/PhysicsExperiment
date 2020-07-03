from numpy import asarray, linspace, sqrt, polyfit, poly1d, float, sqrt
from mpl_toolkits.axisartist.axislines import SubplotZero
from sympy import symbols, diff, sympify, integrate
from matplotlib import pyplot as plt
from scipy.optimize import fsolve
from scipy.misc import derivative as dx


class Calculus:  # 一元微积分计算
	"""
	定积分
	@param
	expression: 要积分的显函数表达式，仅允许有x一个未知量，无需dx和积分号
	lower: 积分下限
	upper: 积分上限
	@return
	积分结果
	"""
	@staticmethod
	def definite_integral(expression, lower, upper):
		x = symbols('x')
		return integrate(expression, (x, lower, upper))

	'''
	不定积分
	@param
	expression: 要积分的显函数表达式，仅允许有x一个未知量，无需dx和积分号
	@return
	积分结果，一个sympy数学表达式
	'''
	@staticmethod
	def indefinite_integral(expression):
		x = symbols('x')
		return integrate(expression, x)

	'''
	求导函数
	@param
	expression: 要求导的显函数表达式，仅允许有x一个未知量
	@return
	求导结果，一个sympy数学表达式
	'''
	@staticmethod
	def derivative(expression):
		x = symbols('x')
		return diff(expression, x)

	'''
	求导数
	@param
	func: 要求导的函数
	x: 在x点导数值
	@return
	求导结果
	'''
	@staticmethod
	def derivative_point(func, x):
		return dx(func, x, 1e-7)

	'''
	代入sympy表达式
	@param
	expression: 要积分的显函数表达式，仅允许有x一个未知量，无需dx和积分号
	value: 要代入的x值
	@return
	计算结果
	'''
	@staticmethod
	def f(expression, value):
		x = symbols('x')
		return sympify(expression).evalf(subs={x: value})


class InstrumentError:  # 仪器误差限计算
	"""
	直流电位差计仪器误差限
	@param
	a：误差系数（a%中的a）
	Ux：电位差计的测量值
	U0：有效量程的基准值（该量程中最大的10的整数幂）
	@return
	一个数字，代表仪器误差限
	"""
	@staticmethod
	def dc_potentiometer(a, Ux, U0):
		return 0.01 * a * (Ux + U0 / 10)

	'''
	直流电桥仪器误差限
	@param
	a：误差系数（a%中的a）
	Rx：直流电桥的测量值
	R0：有效量程的基准值（该量程中最大的10的整数幂）
	@return
	一个数字，代表仪器误差限
	'''
	@staticmethod
	def dc_bridge(a, Rx, R0):
		return 0.01 * a * (Rx + R0 / 10)

	'''
	数字仪表误差限
	这里有两种表示形式，调用时务必明确表示
	调用方式：
	digital_instrument(a=1, Nx=10.0, b=2, Nm=20.0)
	digital_instrument(a=1, Nx=10.0, n=3, scale=0.01)
	'''
	@staticmethod
	def digital_instrument(a, Nx, **kwargs):
		if kwargs.__contains__('b') and kwargs.__contains__('Nm'):
			b = kwargs['b']
			Nm = kwargs['Nm']
			return 0.01 * (a * Nx + b * Nm)
		elif kwargs.__contains__('n') and kwargs.__contains__('scale'):
			n = kwargs['n']
			scale = kwargs['scale']
			return 0.01 * a * Nx + n * scale

	'''
	钢直尺误差限
	@param
	length: 测量长度
	@return
	一个数字，代表最大限度
	'''
	@staticmethod
	def steel_ruler(length):
		if length < 0:
			return None
		if length <= 300:
			return 0.10
		elif length <= 500:
			return 0.15
		elif length <= 1000:
			return 0.20
		elif length <= 1500:
			return 0.27
		elif length <= 2000:
			return 0.35
		else:
			return None

	'''
	钢卷尺误差限
	@param
	level: 等级
	length: 测量长度
	@return
	一个数字，代表最大限度
	'''
	@staticmethod
	def steel_tape(level, length):
		if level == 1:
			return 0.1 + 0.001 * length
		elif level == 2:
			return 0.3 + 0.002 * length
		else:
			return None

	'''
	游标卡尺误差限
	@param
	division: 分度值
	length: 测量长度
	@return
	一个数字，代表最大限度
	'''
	@staticmethod
	def caliper(division, length):
		if division == 0.02:
			if length <= 150:
				return 0.02
			elif length <= 200:
				return 0.03
			elif length <= 300:
				return 0.04
			elif length <= 500:
				return 0.05
			elif length <= 1000:
				return 0.07
			else:
				return None
		elif division == 0.05:
			if length <= 150:
				return 0.05
			elif length <= 200:
				return 0.05
			elif length <= 300:
				return 0.08
			elif length <= 500:
				return 0.08
			elif length <= 1000:
				return 0.1
			else:
				return None
		elif division == 0.1:
			if length <= 500:
				return 0.10
			elif length <= 1000:
				return 0.15
			else:
				return None
		else:
			return None

	'''
	螺旋测微器误差限
	@param
	length: 测量长度
	@return
	一个数字，代表最大限度
	'''
	@staticmethod
	def micrometer(length):
		if length <= 50:
			return 0.004
		elif length <= 100:
			return 0.005
		elif length <= 150:
			return 0.006
		elif length <= 200:
			return 0.007
		else:
			return None

	'''
	求磁电式电表的仪器误差限
	@param
	a：误差系数（a%中的a）
	Nm：仪器量程
	@return
	一个数字，代表仪器误差限
	'''
	@staticmethod
	def electromagnetic_instrument(a, Nm):
		return a * Nm / 100

	'''
	求直流电阻器的仪器误差限
	@param
	r20：20摄氏度时的电阻
	a：一次电阻系数α
	b：二次电阻系数β
	t：仪器的使用温度
	@return
	一个数字，代表仪器误差限
	'''
	@staticmethod
	def dc_resistor(r20, a, b, t):
		return r20 * (1 + a * (t - 20) + b * ((t - 20) ** 2))

	'''
	求电阻箱的仪器误差限
	@param
	a：每一级误差系数（a%中的a，一维数组）
	r：每一级的示数（一维数组，和a中的数据一一对应）
	r0：残余电阻
	@return
	一个数字，代表仪器误差限
	'''
	@staticmethod
	def resistance_box(a, r, r0):
		sum1 = 0
		for i in range(len(a)):
			sum1 += a[i] * r[i] / 100
		return sum1 + r0

	'''
	时间测量误差限
	@param
	无
	@return
	一个数字，代表最大限度
	'''
	@staticmethod
	def time():
		return 5.8 * 10 ** -6 + 0.01

	'''
	全浸式水银温度计误差限
	@param
	t: 测量温度
	division: 分度值
	@return
	一个数字，代表最大限度
	'''
	@staticmethod
	def full_immersion_mercury_thermometer(t, division):
		if -30 <= t <= 100:
			if division == 0.1:
				return 0.2
			elif division == 0.2:
				return 0.3
			elif division == 0.5:
				return 0.5
			elif division == 1:
				return 1
		elif 100 < t <= 200:
			if division == 0.1:
				return 0.4
			elif division == 0.2:
				return 0.4
			elif division == 0.5:
				return 1.0
			elif division == 1:
				return 1.5

	'''
	局浸式水银温度计误差限
	@param
	t: 测量温度
	division: 分度值
	@return
	一个数字，代表最大限度
	'''
	@staticmethod
	def local_immersion_mercury_thermometer(t, division):
		if -30 <= t <= 100:
			if division == 0.5:
				return 1.0
			elif division == 1:
				return 1.5
		elif 100 < t <= 200:
			if division == 0.5:
				return 1.5
			elif division == 1:
				return 2

	'''
	工作用铂铑-铂热电偶温度计误差限
	@param
	t: 测量温度
	level: 等级
	@return
	一个数字，代表最大限度
	'''
	@staticmethod
	def Pt_Rh_couple(level, t):
		if level == 1:
			if 0 <= t <= 1100:
				return 1
			elif 1100 <= t <= 1600:
				return 1 + (t - 1100) * 0.003
		elif level == 2:
			if 0 <= t <= 600:
				return 1.5
			elif 600 <= t <= 1600:
				return 0.0025 * t

	'''
	工业铂电阻温度计误差限
	@param
	t: 测量温度
	level: 等级
	@return
	一个数字，代表最大限度
	'''
	@staticmethod
	def Pt_resistance(level, t):
		if level == 'A':
			return 0.15 + 0.002 * abs(t)
		elif level == 'B':
			return 0.30 + 0.005 * abs(t)


class Simplified:   # 长度时间测量误差限简化版，直接继承就行
	def __init__(self):
		self.steel_ruler = 0.5
		self.steel_tape = 0.5
		self.micrometer = 0.005
		self.time = 0.01

	@staticmethod
	def caliper(division):
		if division == 0.1:
			return 0.1
		elif division == 0.05:
			return 0.05
		elif division == 0.02:
			return 0.02
		else:
			return None


class Fitting:  # 拟合计算
	"""
	一元线性回归拟合
	@param
	X, Y: 两个等长一维数组，对应位置元素组成散点的坐标
	show_plot: 默认为False, 如果为True则调用时新开一个窗口显示图象
	@return
	(a, b, r): 返回一个三元组，按照顺序分别为回归直线的斜率a，截距b以及线性相关程度r
	"""
	@staticmethod
	def linear(X, Y, show_plot=False):
		X, Y = asarray(X), asarray(Y)
		if len(X.shape) != 1 or len(Y.shape) != 1:
			raise ValueError("The dimension of X and Y must be both 1")
		if X.shape[0] != Y.shape[0]:
			raise ValueError("The length of X and Y must be equal")
		X = X.astype(float)
		Y = Y.astype(float)
		XY = X * Y
		XY_mean = XY.mean()
		X_mean = X.mean()
		Y_mean = Y.mean()
		X2_mean = (X ** 2).mean()

		a_hat = (XY_mean - X_mean * Y_mean) / (X2_mean - X_mean ** 2)
		b_hat = Y_mean - a_hat * X_mean

		r = (((X - X_mean) * (Y - Y_mean)).sum()) / (
				sqrt(((X - X_mean) ** 2).sum()) * sqrt(((Y - Y_mean) ** 2).sum()))
		if show_plot:
			plt.figure().canvas.set_window_title("一元线性回归拟合")
			plt.scatter(X, Y, marker='+')
			plt.plot(X, a_hat * X + b_hat, c='red', linewidth=1)
			plt.show()
		return [a_hat, b_hat, r]

	'''
	多项式拟合
	@param
	X, Y: 两个等长一维数组，对应位置元素组成散点的坐标
	deg: 正整数，表示几次多项式去拟合
	show_plot: 默认为False, 如果为True则调用时新开一个窗口显示图象，为散点列和拟合的直线
	@return
	一个二元组(p, f)，其中p为系数数组，顺序为从常数项到高次项, f为一个函数(Callable)，即拟合的多项式
	'''
	@staticmethod
	def poly(x, y, deg, show_plot=False):
		x, y = asarray(x), asarray(y)
		if len(x.shape) != 1 or len(y.shape) != 1:
			raise ValueError("The dimension of X and Y must be both 1")
		if x.shape[0] != y.shape[0]:
			raise ValueError("The length of X and Y must be equal")
		p = polyfit(x, y, deg)  # 系数
		f = poly1d(p)
		if show_plot:
			x_left, x_right = min(x), max(x)
			xx = linspace(x_left, x_right)
			# plt.clf()
			plt.figure().canvas.set_window_title("多项式拟合")
			plt.scatter(x, y, marker='+')
			plt.plot(xx, f(xx), c='red', linewidth=1)
			plt.show()
		return p, f

	'''
	逐差法拟合
	@param
	X, Y: 两个等长一维数组，对应位置元素组成散点的坐标
	@return
	一个二元组(a, b)，代表 y = a + bx中的a, b
	'''
	@staticmethod
	def successive_difference(x, y):
		if len(x) != len(y):
			print("X数组与Y数组长度不相等，无法进行逐差法计算")
			return
		if len(x) & 1 != 0:
			x.pop()
			y.pop()
		n = int(len(x) / 2)
		sumd = 0
		for i in range(len(x)):
			if i == n:
				break
			sumd += (y[n + i] - y[i]) / (x[n + i] - x[i])
		b = (1 / n) * sumd
		sumx = 0
		sumy = 0
		for i in range(len(x)):
			sumx += x[i]
			sumy += y[i]
		a = (1 / (2 * n)) * (sumy - b * sumx)
		return a, b

	'''
	绘制函数图像
	@param
	section：二元组[a,b]，表示画函数图像的区间
	func：自己定义的函数，只有x一个自变量
	@return
	无，会弹出绘图窗口，绘制函数图像
	'''
	@staticmethod
	def plot_func(section, func):
		fig = plt.figure(1, (10, 6))
		ax = SubplotZero(fig, 1, 1, 1)
		fig.add_subplot(ax)
		x0 = linspace(section[0], section[1], 1000)
		y0 = []
		for i in range(1000):
			y0.append(func(x0[i]))
		ax.axis["xzero"].set_visible(True)
		ax.axis["xzero"].label.set_color('green')
		ax.axis["yzero"].set_visible(True)
		ax.axis["yzero"].label.set_color('green')
		plt.plot(x0, y0, 'r-', color='b')
		plt.show()

	'''
	绘制函数图像（字符串函数）
	@param
	section：二元组[a,b]，表示画函数图像的区间
	expression：字符串形式的函数表达式，只有x一个自变量
	@return
	无，会弹出绘图窗口，绘制函数图像
	'''
	@staticmethod
	def plot_func_str(section, expression):
		fig = plt.figure(1, (10, 6))
		ax = SubplotZero(fig, 1, 1, 1)
		fig.add_subplot(ax)
		x0 = linspace(section[0], section[1], 1000)
		x = symbols('x')
		y0 = []
		for i in range(1000):
			y0.append(sympify(expression).evalf(subs={x: x0[i]}))
		ax.axis["xzero"].set_visible(True)
		ax.axis["xzero"].label.set_color('green')
		ax.axis["yzero"].set_visible(True)
		ax.axis["yzero"].label.set_color('green')
		plt.plot(x0, y0, 'r-', color='b')
		plt.show()

	'''
	求函数图象的交点
	使用方法：调用此函数会先自动绘制两个函数在给定范围的图象，用鼠标点击交点则屏幕上可以显示出交点坐标
	@param
	func1, func2: 两个一元函数, 为python中可以调用的函数，参数为一个自变量，返回值为函数值
	range: 一个二元组，表示x的范围
	@return
	void
	Usage: intersection(lambda x : x**2, lambda x : x, [-1, 4])
	'''
	@staticmethod
	def intersection(func1, func2, section):
		import warnings
		warnings.filterwarnings('ignore')
		fdif = lambda x: func1(x) - func2(x)

		def onPress(event):
			if event.button == 1:
				px = event.xdata
				eqroot = fsolve(fdif, px)
				rooty = func1(eqroot)
				plt.ion()
				plt.scatter(eqroot, rooty, marker='+', c='red')
				plt.text(eqroot, rooty, ("x=%.3f\ny=%.3f" % (eqroot, rooty)))
				plt.ioff()
				plt.show()

		x = linspace(section[0], section[1])
		fig = plt.figure()
		plt.rcParams['font.sans-serif'] = ['SimHei', 'SimSun'] # 使matplotlib支持中文字体
		plt.rcParams['axes.unicode_minus'] = False # 正常显示负号
		plt.plot(x, func1(x))
		plt.plot(x, func2(x))
		plt.title("请用鼠标点击函数交点")
		fig.canvas.set_window_title("函数图象交点查看器")
		fig.canvas.mpl_connect('button_press_event', onPress)
		plt.show()


class Method:
	"""
	求最大公约数
	@param
	x, y：需要求最大公约数的两个数
	@return
	最大公约数的值
	"""
	@staticmethod
	def gcd(x, y):
		z = min(x, y)
		while x % z != 0 or y % z != 0:
			z = z - 1
		return z

	"""
	求最小公倍数
	@param
	x, y：需要求最小公倍数的两个数
	@return
	最小公倍数的值
	"""
	@staticmethod
	def lcm(x, y):
		z = max(x, y)
		while z % x != 0 or z % y != 0:
			z = z + 1
		return z

	'''
	求平均数数
	@param
	a：需要求平均数的列表
	@return
	列别的平均数的值
	'''

	@staticmethod
	def average(a):
		sum1 = 0
		for i in a:
			sum1 = sum1 + i
		av = sum1 / (len(a))
		return av

	'''
	求方差
	@param
	nums：需要求方差的列表
	@return
	方差的值
	'''
	@staticmethod
	def variance(nums):
		av = Method.average(nums)
		sum1 = 0
		for i in nums:
			sum1 = (i - av) ** 2 + sum1
		s = sum1 / len(nums)
		return s

	'''
	求A类不确定度
	@param
	a：一维数据数组
	@return
	一个数字，代表数组中所有数据的A类不确定度
	'''
	@staticmethod
	def a_uncertainty(a):
		abu = (Method.variance(a) / (len(a) - 1)) ** (1 / 2)
		return abu

