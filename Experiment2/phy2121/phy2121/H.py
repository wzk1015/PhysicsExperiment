import xlrd
import shutil
import os
from numpy import sqrt, abs,sin,cos,tan
import sys
sys.path.append('../..') 
from GeneralMethod.PyCalcLib import Method

from GeneralMethod.Report import Report



class exp2121:
    # 需往实验报告中填的空的key，这些key在Word模板中以#号包含，例如#1#, #delta_d#, #final#

    PREVIEW_FILENAME = "Preview.pdf"
    DATA_SHEET_FILENAME = "data.xlsx"
    REPORT_TEMPLATE_FILENAME = "H_empty.docx"
    REPORT_OUTPUT_FILENAME = "H_out.docx"

    def __init__(self):
        self.data = {} # 存放实验中的各个物理量
        self.uncertainty = {} # 存放物理量的不确定度
        self.report_data = {} # 存放需要填入实验报告的
        print("2121 氢原子光谱实验\n1. 实验预习\n2. 数据处理")
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
            print("现在开始数据处理")
            print("即将打开数据输入文件......")
            # 打开数据输入文件
            Method.start_file(self.DATA_SHEET_FILENAME)
            input("输入数据完成后请保存并关闭excel文件，然后按回车键继续")
            # 从excel中读取数据
            self.input_data("./"+self.DATA_SHEET_FILENAME) # './' is necessary when running this file, but should be removed if run main.py
            print("数据读入完毕，处理中......")
            # 计算物理量
            self.calc_data_d()
            # 计算不确定度
            self.calc_uncertainty_d()
            #下一问
            self.calc_data_Rh()
            self.calc_uncertainty_Rh()
            self.calc_data_3()
            print("正在生成实验报告......")
            # 生成实验报告
            self.fill_report()
            print("实验报告生成完毕，正在打开......")
            Method.start_file(self.REPORT_OUTPUT_FILENAME)
            print("Done!")

    '''
    从excel表格中读取数据
    @param filename: 输入excel的文件名
    @return none
    '''
    def input_data(self, filename):
       
        wb = xlrd.open_workbook(filename) 
        ws = wb.sheet_by_name('Ud')
        # 从excel中读取数据，首先是第一个实验：d及其不确定度的计算
        
        list_d_a1 = []
        list_d_b1 = []
        list_d_a2 = []
        list_d_b2 = []
        for row in range(2,7):
            list_d_a1.append(float(self.Degree(ws.cell_value(row, 1)))) # 从excel取出来的数据，加个类型转换靠谱一点
            list_d_b1.append(float(self.Degree(ws.cell_value(row, 2))))
            list_d_a2.append(float(self.Degree(ws.cell_value(row, 3))))
            list_d_b2.append(float(self.Degree(ws.cell_value(row, 4))))
        a1_2=float(self.Degree(ws.cell_value(10, 1)))
        b1_2=float(self.Degree(ws.cell_value(10, 2)))
        a2_2=float(self.Degree(ws.cell_value(10, 3)))
        b2_2=float(self.Degree(ws.cell_value(10, 4)))
                    
        self.data['list_d_a1'] = list_d_a1 # 存储从表格中读入的数据
        self.data['list_d_b1'] = list_d_b1
        self.data['list_d_a2'] = list_d_a2
        self.data['list_d_b2'] = list_d_b2
        self.data['a1_2'] =  a1_2
        self.data['b1_2'] =  b1_2
        self.data['a2_2'] =  a2_2
        self.data['b2_2'] =  b2_2
        
        #接下来是氢原子光谱三种光★
        ws = wb.sheet_by_name('three')
        list_r_a1 = []
        list_r_b1 = []
        list_r_a2 = []
        list_r_b2 = []
        for row in range(2, 7):
            list_r_a1.append(float(self.Degree(ws.cell_value(row, 1)))) # 从excel取出来的数据，加个类型转换靠谱一点
            list_r_b1.append(float(self.Degree(ws.cell_value(row, 2))))
            list_r_a2.append(float(self.Degree(ws.cell_value(row, 3))))
            list_r_b2.append(float(self.Degree(ws.cell_value(row, 4))))
                    
        self.data['list_r_a1'] = list_r_a1 # 存储从表格中读入的数据
        self.data['list_r_b1'] = list_r_b1
        self.data['list_r_a2'] = list_r_a2
        self.data['list_r_b2'] = list_r_b2
        #蓝
        list_b_a1 = []
        list_b_b1 = []
        list_b_a2 = []
        list_b_b2 = []
        for row in range(10, 15):
            list_b_a1.append(float(self.Degree(ws.cell_value(row, 1)))) # 从excel取出来的数据，加个类型转换靠谱一点
            list_b_b1.append(float(self.Degree(ws.cell_value(row, 2))))
            list_b_a2.append(float(self.Degree(ws.cell_value(row, 3))))
            list_b_b2.append(float(self.Degree(ws.cell_value(row, 4))))
                    
        self.data['list_b_a1'] = list_b_a1 # 存储从表格中读入的数据
        self.data['list_b_b1'] = list_b_b1
        self.data['list_b_a2'] = list_b_a2
        self.data['list_b_b2'] = list_b_b2
        #紫
        list_p_a1 = []
        list_p_b1 = []
        list_p_a2 = []
        list_p_b2 = []
        for row in range(18, 23):
            list_p_a1.append(float(self.Degree(ws.cell_value(row, 1)))) # 从excel取出来的数据，加个类型转换靠谱一点
            list_p_b1.append(float(self.Degree(ws.cell_value(row, 2))))
            list_p_a2.append(float(self.Degree(ws.cell_value(row, 3))))
            list_p_b2.append(float(self.Degree(ws.cell_value(row, 4))))
                    
        self.data['list_p_a1'] = list_p_a1 # 存储从表格中读入的数据
        self.data['list_p_b1'] = list_p_b1
        self.data['list_p_a2'] = list_p_a2
        self.data['list_p_b2'] = list_p_b2
    
    '''
    进行数据处理
    我给2121下跪了
    注意：若需计算的物理量较多，建议对计算过程复杂的物理量单独封装函数.
    对于实验中重要的数据，采用dict对象self.data存储，方便其他函数共用数据
    '''
    
    def simple_2xita(self,list1,list2,list3,list4,i):
        list11=[0,0,0,0,0]
        list21=[0,0,0,0,0]
        for j in range(5):
            if list1[j]<list3[j]:
                list11[j]=list1[j]+360
            else:
                list11[j]=list1[j]
            if list2[j]<list4[j]:
                list21[j]=list2[j]+360
            else:
                list21[j]=list2[j]
            
        return 3.14*((list11[i]-list3[i])+list21[i]-list4[i])/360
        
        
    def calc_data_d(self):
        list_d_2xita1=[]
        list_d_2xita1_print=[]
        for i in range(5):
            list_d_2xita1.append(self.simple_2xita(self.data['list_d_a1'],self.data['list_d_b1'],self.data['list_d_a2'],self.data['list_d_b2'],i))
            list_d_2xita1_print.append(180*list_d_2xita1[i]/3.14)
        self.data['list_d_2xita1'] = list_d_2xita1
        self.data['list_d_2xita1_print'] = list_d_2xita1_print
        d_2xita2 = ((self.data['a1_2']-self.data['a2_2'])+self.data['b1_2']-self.data['b2_2'])/2
        self.data['d_2xita2'] =  d_2xita2
        d_2xita1_average=Method.average(list_d_2xita1)
        d_xita1_average=   d_2xita1_average/2
        self.data['d_2xita1_average'] = d_2xita1_average
        self.data['d_xita1_average'] = d_xita1_average
        d1=589.3/sin(d_xita1_average)
        self.data['d1'] = d1 / 1000
        d2=2*589.3/sin(3.14*d_2xita2/360)
        self.data['d2'] = d2 / 1000
        
    def calc_data_Rh(self):
        #三种光要算Rh啦（r为红，b为蓝，p为紫）
        list_r_2xita1=[]
        list_r_2xita1_print=[]
        for i in range(5):
            list_r_2xita1.append(self.simple_2xita(self.data['list_r_a1'],self.data['list_r_b1'],self.data['list_r_a2'],self.data['list_r_b2'],i))
            list_r_2xita1_print.append(180*list_r_2xita1[i]/3.14)
        self.data['list_r_2xita1'] = list_r_2xita1
        self.data['list_r_2xita1_print'] = list_r_2xita1_print
        r_2xita1_average=Method.average(list_r_2xita1)
        r_xita1_average=   r_2xita1_average/2
        self.data['r_2xita1_average'] = r_2xita1_average
        self.data['r_xita1_average'] = r_xita1_average
        lambdaR=self.uncertainty['d_av']*sin(r_xita1_average)
        r_RH=36/(5*lambdaR)
        self.data['r_2xita1_average'] = r_2xita1_average
        self.data['r_xita1_average'] = r_xita1_average
        self.data['r_RH'] =  r_RH
        self.data['lambdaR'] =  lambdaR
        
        list_b_2xita1=[]
        list_b_2xita1_print=[]
        for i in range(5):
            list_b_2xita1.append(self.simple_2xita(self.data['list_b_a1'],self.data['list_b_b1'],self.data['list_b_a2'],self.data['list_b_b2'],i))
            list_b_2xita1_print.append(180*list_b_2xita1[i]/3.14)
        self.data['list_b_2xita1'] = list_b_2xita1
        self.data['list_b_2xita1_print'] = list_b_2xita1_print
        b_2xita1_average=Method.average(list_b_2xita1)
        b_xita1_average=   b_2xita1_average/2
        self.data['b_2xita1_average'] = b_2xita1_average
        self.data['b_xita1_average'] = b_xita1_average
        lambdaB=self.uncertainty['d_av']*sin(b_xita1_average)
        b_RH=16/(3*lambdaB)
        self.data['b_2xita1_average'] = b_2xita1_average
        self.data['b_xita1_average'] = b_xita1_average
        self.data['b_RH'] =  b_RH
        self.data['lambdaB'] =  lambdaB
        #purple
        list_p_2xita1=[]
        list_p_2xita1_print=[]
        for i in range(5):
            list_p_2xita1.append(self.simple_2xita(self.data['list_p_a1'],self.data['list_p_b1'],self.data['list_p_a2'],self.data['list_p_b2'],i))
            list_p_2xita1_print.append(180*list_p_2xita1[i]/3.14)
        self.data['list_p_2xita1'] = list_p_2xita1
        self.data['list_p_2xita1_print'] = list_p_2xita1_print
        p_2xita1_average=Method.average(list_p_2xita1)
        p_xita1_average=   p_2xita1_average/2
        self.data['p_2xita1_average'] = p_2xita1_average
        self.data['p_xita1_average'] = p_xita1_average
        lambdaP=self.uncertainty['d_av']*sin(p_xita1_average)
        p_RH=100/(21*lambdaP)
        self.data['p_2xita1_average'] = p_2xita1_average
        self.data['p_xita1_average'] = p_xita1_average
        self.data['p_RH'] =  p_RH
        self.data['lambdaP'] =  lambdaP
        
    def calc_data_3(self):
        N=2.2/self.uncertainty['d_av']
        R1=N
        R2=2*N
        DD1=1/(self.uncertainty['d_av']*cos(self.data['d_xita1_average']))
        DD2=2/(self.uncertainty['d_av']*cos(3.14*self.data['d_2xita2']/360))
        DELTA_lambda1=(589.3/R1)*1e-4
        DELTA_lambda2=(589.3/R2)*1e-4
        self.data['N'] =  N
        self.data['R1'] =  R1
        self.data['R2'] =  R2
        self.data['DD1'] =  DD1
        self.data['DD2'] =  DD2
        self.data['DELTA_lambda1'] =  DELTA_lambda1
        self.data['DELTA_lambda2'] =  DELTA_lambda2
        
        
    '''
    计算所有的不确定度
    '''
    # 对于数据处理简单的实验，可以根据此格式，先计算数据再算不确定度，若数据处理复杂也可每计算一个物理量就算一次不确定度
    def calc_uncertainty_d(self):
        # 计算光程差d的a,b及总不确定度
        Ua_2xita1 = Method.a_uncertainty(self.data['list_d_2xita1']) # 这里容易写错，一定要用原始数据的数组
        Ub_2xita1 = 1.679e-4
        U_xita1 = sqrt(Ub_2xita1 ** 2 + Ua_2xita1 ** 2)/2
        Ud1=abs(self.data['d1']*U_xita1/tan(self.data['d_xita1_average']) )
        Ub_2xita2 = 1.679e-4
        Ud2=abs(self.data['d2']*Ub_2xita2/(2*tan(3.14*self.data['d_2xita2']/360)))
        self.uncertainty.update({"Ua_2xita1":Ua_2xita1, "Ub_2xita1":Ub_2xita1,"Ua_xita1":Ua_2xita1/2, "Ub_xita1":Ub_2xita1/2, "U_xita1":U_xita1,"Ud1":Ud1, "Ub_2xita2":Ub_2xita2,"Ub_xita2":Ub_2xita2/2, "Ud2":Ud2})
        d1=self.data['d1']
        d2=self.data['d2']
        d_av=(d1/(Ud1**2)+d2/(Ud2**2))/(1/(Ud1**2)+1/(Ud2**2))
        Ud_av=sqrt((Ud1**2)*(Ud2**2)/(Ud1**2)+(Ud2**2))
        self.uncertainty.update({"d_av":d_av,"Ud_av":Ud_av})
        self.data['final1'] = Method.final_expression(d_av, Ud_av) 
        
        
    def calc_uncertainty_Rh(self):
        Ua_2xita1_r = Method.a_uncertainty(self.data['list_r_2xita1']) # 这里容易写错，一定要用原始数据的数组
        Ub_2xita1_r = 1.679e-4
        Ua_xita1_r = Ua_2xita1_r / 2 
        Ub_xita1_r = Ub_2xita1_r / 2
        U_xita1_r = sqrt(Ub_xita1_r ** 2 + Ua_xita1_r ** 2)
        self.uncertainty.update({"Ua_xita1_r":Ua_xita1_r, "Ub_xita1_r":Ub_xita1_r, "U_xita1_r":U_xita1_r,"Ua_2xita1_r":Ua_2xita1_r, "Ub_2xita1_r":Ub_2xita1_r})
        U_r_RH=self.data['r_RH'] *sqrt((self.uncertainty['Ud_av']/self.uncertainty['d_av'])**2+(U_xita1_r/tan(self.data['r_xita1_average'])**2))
        self.uncertainty.update({"U_r_RH":U_r_RH})
        self.data['final2_r'] = Method.final_expression(self.data['r_RH']*1e6, U_r_RH*1e4)
        
        Ua_2xita1_b = Method.a_uncertainty(self.data['list_b_2xita1']) # 这里容易写错，一定要用原始数据的数组
        Ub_2xita1_b = 1.679e-4
        Ua_xita1_b = Ua_2xita1_b / 2 
        Ub_xita1_b = Ub_2xita1_b / 2
        U_xita1_b = sqrt(Ub_xita1_b ** 2 + Ua_xita1_b ** 2)
        self.uncertainty.update({"Ua_xita1_b":Ua_xita1_b, "Ub_xita1_b":Ub_xita1_b, "U_xita1_b":U_xita1_b,"Ua_2xita1_b":Ua_2xita1_b, "Ub_2xita1_b":Ub_2xita1_b})
        U_b_RH=self.data['b_RH'] *sqrt((self.uncertainty['Ud_av']/self.uncertainty['d_av'])**2+(U_xita1_b/tan(self.data['b_xita1_average'])**2))
        self.uncertainty.update({"U_b_RH":U_b_RH})
        self.data['final2_b'] = Method.final_expression(self.data['b_RH']*1e6, U_b_RH*1e4)
        
        Ua_2xita1_p = Method.a_uncertainty(self.data['list_p_2xita1']) # 这里容易写错，一定要用原始数据的数组
        Ub_2xita1_p = 1.679e-4
        Ua_xita1_p = Ua_2xita1_p / 2 
        Ub_xita1_p = Ub_2xita1_p / 2
        U_xita1_p = sqrt(Ub_xita1_p ** 2 + Ua_xita1_p ** 2)
        self.uncertainty.update({"Ua_xita1_p":Ua_xita1_p, "Ub_xita1_p":Ub_xita1_p, "U_xita1_p":U_xita1_p,"Ua_2xita1_p":Ua_2xita1_p, "Ub_2xita1_p":Ub_2xita1_p})
        U_p_RH=self.data['p_RH'] *sqrt((self.uncertainty['Ud_av']/self.uncertainty['d_av'])**2+(U_xita1_p/tan(self.data['p_xita1_average'])**2))
        self.uncertainty.update({"U_p_RH":U_p_RH})
        self.data['final2_p'] = Method.final_expression(self.data['p_RH']*1e6, U_p_RH*1e4)
        
        RH=(self.simple_Rh(self.data['r_RH'],U_r_RH)+self.simple_Rh(self.data['b_RH'],U_b_RH)+self.simple_Rh(self.data['p_RH'],U_p_RH))/((1/U_r_RH**2)+(1/U_b_RH**2)+(1/U_p_RH**2))
        self.data['RH']=RH
        U_RH=sqrt(1/((1/U_r_RH**2)+(1/U_b_RH**2)+(1/U_p_RH**2)))
        self.data['U_RH']=U_RH
        self.data['final2'] = Method.final_expression(RH*1e6,U_RH*1e4) 
    def simple_Rh(self,a,ua):
        return a/(ua**2)
    
    '''
    填充实验报告
    调用ReportWriter类，将数据填入Word文档格式的实验报告中
    '''

    def Degree(self,degree):
        degree_list = degree.split(' ')
        if len(degree_list) == 1:
            return eval(degree)
        elif len(degree_list) == 2:
            return float(eval(degree_list[0])) + float(eval(degree_list[1]) / 60.0)
        else:
            return False
        
    def fill_report(self):
        # 表格：原始数据d
        for i, d_i in enumerate(self.data['list_d_a1']):
            self.report_data['d_a1_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_d_a2']):
            self.report_data['d_a2_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_d_b1']):
            self.report_data['d_b1_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_d_b2']):
            self.report_data['d_b2_'+str(i + 1)] = "%.5f" % (d_i)# 一定都是字符串类型
        for i, d_i in enumerate(self.data['list_d_2xita1_print']):
            self.report_data['d_2xita1_'+str(i + 1)] = "%.5f" % (d_i)
        self.report_data['a1_2']="%.5f" % self.data['a1_2']
        self.report_data['b1_2']="%.5f" % self.data['b1_2']
        self.report_data['a2_2']="%.5f" % self.data['a2_2']
        self.report_data['b2_2']="%.5f" % self.data['b2_2']
        self.report_data['d_2xita2']="%.5f" % self.data['d_2xita2']
        self.report_data['d_2xita1_average']=self.data['d_2xita1_average']
        self.report_data['d_xita1_average']=self.data['d_xita1_average']
        self.report_data['d1']=self.data['d1']
        self.report_data['d2']=self.data['d2']
        self.report_data['Ua_2xita1']=self.uncertainty['Ua_2xita1']
        self.report_data['Ub_2xita1']=self.uncertainty['Ub_2xita1']
        self.report_data['U_xita1']=self.uncertainty['U_xita1']
        self.report_data['Ud1']=self.uncertainty['Ud1']
        self.report_data['Ub_2xita2']=self.uncertainty['Ub_2xita2']
        self.report_data['Ub_xita2']=self.uncertainty['Ub_xita2']
        self.report_data['Ud2']=self.uncertainty['Ud2']
        self.report_data['d_av']=self.uncertainty['d_av']
        self.report_data['Ud_av']=self.uncertainty['Ud_av']
        self.report_data['final1']=self.data['final1']
        #红光
        for i, d_i in enumerate(self.data['list_r_a1']):
            self.report_data['r_a1_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_r_a2']):
            self.report_data['r_a2_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_r_b1']):
            self.report_data['r_b1_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_r_b2']):
            self.report_data['r_b2_'+str(i + 1)] = "%.5f" % (d_i)# 一定都是字符串类型
        for i, d_i in enumerate(self.data['list_r_2xita1_print']):
            self.report_data['r_2xita1_'+str(i + 1)] = "%.5f" % (d_i)
            
        self.report_data['r_2xita1_average']=self.data['r_2xita1_average']
        self.report_data['lambdaR']=self.data['lambdaR']*1000
        self.report_data['r_RH']=self.data['r_RH']*1e6
        self.report_data['Ua_2xita1_r']=self.uncertainty['Ua_2xita1_r']
        self.report_data['Ub_2xita1_r']=self.uncertainty['Ub_2xita1_r']
        self.report_data['U_xita1_r']=self.uncertainty['U_xita1_r']
        self.report_data['U_r_RH']=self.uncertainty['U_r_RH']*1e4
        self.report_data['final2_r']=self.data['final2_r']
        #蓝光   
        for i, d_i in enumerate(self.data['list_b_a1']):
            self.report_data['b_a1_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_b_a2']):
            self.report_data['b_a2_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_b_b1']):
            self.report_data['b_b1_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_b_b2']):
            self.report_data['b_b2_'+str(i + 1)] = "%.5f" % (d_i)# 一定都是字符串类型
        for i, d_i in enumerate(self.data['list_b_2xita1_print']):
            self.report_data['b_2xita1_'+str(i + 1)] = "%.5f" % (d_i)
         
        self.report_data['b_2xita1_average']=self.data['b_2xita1_average']
        self.report_data['lambdaB']=self.data['lambdaB']*1000
        self.report_data['b_RH']=self.data['b_RH']*1e6
        self.report_data['Ua_2xita1_b']=self.uncertainty['Ua_2xita1_b']
        self.report_data['Ub_2xita1_b']=self.uncertainty['Ub_2xita1_b']
        self.report_data['U_xita1_b']=self.uncertainty['U_xita1_b']
        self.report_data['U_b_RH']=self.uncertainty['U_b_RH']*1e4
        self.report_data['final2_b']=self.data['final2_b']
        #紫光
        for i, d_i in enumerate(self.data['list_p_a1']):
            self.report_data['p_a1_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_p_a2']):
            self.report_data['p_a2_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_p_b1']):
            self.report_data['p_b1_'+str(i + 1)] = "%.5f" % (d_i)
        for i, d_i in enumerate(self.data['list_p_b2']):
            self.report_data['p_b2_'+str(i + 1)] = "%.5f" % (d_i)# 一定都是字符串类型
        for i, d_i in enumerate(self.data['list_p_2xita1_print']):
            self.report_data['p_2xita1_'+str(i + 1)] = "%.5f" % (d_i)
         
        self.report_data['p_2xita1_average']=self.data['p_2xita1_average']
        self.report_data['lambdaP']=self.data['lambdaP']*1000
        self.report_data['p_RH']=self.data['p_RH']*1e6
        self.report_data['Ua_2xita1_p']=self.uncertainty['Ua_2xita1_p']
        self.report_data['Ub_2xita1_p']=self.uncertainty['Ub_2xita1_p']
        self.report_data['U_xita1_p']=self.uncertainty['U_xita1_p']
        self.report_data['U_p_RH']=self.uncertainty['U_p_RH']*1e4
        self.report_data['final2_p']=self.data['final2_p']
        
        self.report_data['RH']=self.data['RH']*1e6
        self.report_data['U_RH']=self.data['U_RH']*1e4
        self.report_data['final2']=self.data['final2']
        
        self.report_data['N']=self.data['N']
        self.report_data['R1']=self.data['R1']
        self.report_data['R2']=self.data['R2']
        self.report_data['DD1']=self.data['DD1']
        self.report_data['DD2']=self.data['DD2']
        self.report_data['DELTA_lambda1']="%.5f" % self.data['DELTA_lambda1']
        self.report_data['DELTA_lambda2']="%.5f" % self.data['DELTA_lambda2']
        
        
        
        
        
        # 表格：逐差法计算5Δd
        '''for i, dif_d_i in enumerate(self.data['list_dif_d']):
            self.report_data["5d-%d" % (i + 1)] = "%.5f" % (dif_d_i)
        # 最终结果
        self.report_data['final'] = self.data['final']
        # 将各个变量以及不确定度的结果导入实验报告，在实际编写中需根据实验报告的具体要求设定保留几位小数
        self.report_data['N'] = "%d" % self.data['num_N']
        self.report_data['d'] = "%.5f" % self.data['num_d']
        self.report_data['lbd'] = "%.2f" % self.data['num_lbd']
        self.report_data['ua_d'] = "%.5f" % self.data['num_ua_d']
        self.report_data['ub_d'] = "%.5f" % self.data['num_ub_d']
        self.report_data['u_d'] = "%.5f" % self.data['num_u_d']
        self.report_data['u_N'] = "%.5f" % self.data['num_u_N']
        self.report_data['u_lbd_lbd'] = "%.5f" % self.data['num_u_lbd_lbd']
        self.report_data['u_lbd'] = "%.5f" % self.data['num_u_lbd']'''
        # 调用ReportWriter类
        RW = Report()
        RW.load_replace_kw(self.report_data)
        RW.fill_report(self.REPORT_TEMPLATE_FILENAME, self.REPORT_OUTPUT_FILENAME)

if __name__ == '__main__':
    mc = exp2121()
    
