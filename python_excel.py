# _*_ coding:utf-8 _*_
import xlrd
import xlwt
import xlutils
import os
import re
import time
from decimal import Decimal 
from decimal import getcontext

#定义两个全局list，all_info存储每个excel内容，bank_info存储表格内银行名，以便后面根据银行名来确认写excel时输出位置
all_infos = []
bank_info = []

#读取浦发银行流水线，并把相关信息append进入前面定义的两个list
def read_excel_pufa(file_name):
	print('正在读取',file_name,'......')
	wb = xlrd.open_workbook(file_name)
	sh = wb.sheet_by_index(0)
	count = sh.cell(0,1).value #账号
	count_name = sh.cell(1,1).value #账户名 
	if sh.cell(0,0).value != '账号' or sh.cell(1,0).value != '账户名称':
		print(file_name,'银行流水表格式错误！')
		return
	bank_type = file_name.replace('.xls','')[2:]
	l_pufa = [bank_type,count] #银行名和账号放进一个临时list，后面把所有信息一起作为一个list，append进all_info，后面取值方便
	for i in range(sh.nrows):
		if i >= 4:
			if sh.row_values(i)[0] != '':			
				date = sh.row_values(i)[0]	#日期
				
				money_out = sh.row_values(i)[5].replace(' ','') #出
				if len(money_out): #长度大于0时说明有值，转为float类型
					money_out = float(money_out)

				money_in = sh.row_values(i)[6].replace(' ','') #入
				if len(money_in):
					money_in = float(money_in)

				money_now = sh.row_values(i)[7].replace(' ','') #结余
				if len(money_now):
					money_now = float(money_now)

				to_count_name = sh.row_values(i)[9] #对方账号名
				to_count = sh.row_values(i)[8] #对方账户
				beizhu = sh.row_values(i)[10] #备注
	
				l_pufa.append({"count":count,"date":date,"money_out":money_out,"money_in":money_in,"money_now":money_now,"to_count_name":to_count_name,"to_count":to_count,"beizhu":beizhu})
				
				# l.append((count,date,money_out,money_in,money_now,to_count_name,to_count,beizhu))
	global all_infos
	all_infos.append(l_pufa)
	print('读取',file_name,'成功，准备合并数据......')

#读取建设银行流水线，并把相关信息append进入前面定义的两个list
def read_excel_jianhang(file_name):
	print('正在读取',file_name,'......')
	wb = xlrd.open_workbook(file_name)
	sh = wb.sheet_by_index(0)
	if sh.cell(0,0).value != '中国建设银行':
		print(file_name,'银行流水表格式错误！')
		return
	bank_name = sh.cell(3,1).value
	count = sh.cell(4,1).value
	count_name = sh.cell(5,1).value
	bank_type = file_name.replace('.xls','')[2:]
	l_jianhang = [bank_type,count]
	# global bank_info
	# bank_info.append({'bank_name':'建行','count':count})
	for i in range(sh.nrows):
		if i > 8:
			if sh.row_values(i)[0] != '':
				date = sh.row_values(i)[0].replace('-','') #日期格式格式化为20170102
				if isinstance(sh.row_values(i)[4],float):
					money_out = sh.row_values(i)[4]
				else:
					money_out = ''	
				if isinstance(sh.row_values(i)[5],float):
					money_in = sh.row_values(i)[5]
				else:
					money_in = ''		 
				money_now = sh.row_values(i)[6]				
				to_count_name = sh.row_values(i)[8]	
				to_count = sh.row_values(i)[9]
				beizhu = sh.row_values(i)[11]

				# l.append((count,date,money_out,money_in,money_now,to_count_name,to_count,beizhu))
				l_jianhang.append({"count":count,"date":date,"money_out":money_out,"money_in":money_in,"money_now":money_now,"to_count_name":to_count_name,"to_count":to_count,"beizhu":beizhu})
	global all_infos
	all_infos.append(l_jianhang)
	print('读取',file_name,'成功，准备合并数据......')

#读取招商银行流水线，并把相关信息append进入前面定义的两个list
def read_excel_zhaohang(file_name):
	print('正在读取',file_name,'......')
	wb = xlrd.open_workbook(file_name)
	sh = wb.sheet_by_index(0)
	if sh.cell(0,0).value != '交易日':
		print(file_name,'银行流水表格式错误！')
		return
	bank_name = '招商银行'
	count = ''
	count_name = ''
	bank_type = file_name.replace('.xlsx','')[2:]
	l_zhaohang = [bank_type,count]
	# global bank_info
	# bank_info.append({'bank_name':'建行','count':count})
	for i in range(sh.nrows):
		if i > 0:
			if sh.row_values(i)[0] != '':
				date = sh.row_values(i)[0] #日期格式格式化为20170102				
				if isinstance(sh.row_values(i)[1],float):
					money_out = sh.row_values(i)[1]
				else:
					money_out = ''	
				if isinstance(sh.row_values(i)[2],float):
					money_in = sh.row_values(i)[2]
				else:
					money_in = ''		 
				money_now = sh.row_values(i)[3]				
				to_count_name = sh.row_values(i)[5]	
				to_count = sh.row_values(i)[6]
				beizhu = sh.row_values(i)[4]

				# l.append((count,date,money_out,money_in,money_now,to_count_name,to_count,beizhu))
				l_zhaohang.append({"count":count,"date":date,"money_out":money_out,"money_in":money_in,"money_now":money_now,"to_count_name":to_count_name,"to_count":to_count,"beizhu":beizhu})
	global all_infos
	all_infos.append(l_zhaohang)
	print('读取',file_name,'成功，准备合并数据......')

#读取中信银行流水线，并把相关信息append进入前面定义的两个list
def read_excel_zhongxin(file_name):
	print('正在读取',file_name,'......')
	wb = xlrd.open_workbook(file_name)
	sh = wb.sheet_by_index(0)
	if sh.cell(3,0).value != '交易日期':
		print(file_name,'银行流水表格式错误！')
		return
	count = sh.cell(1,3).value #账号
	count_name = sh.cell(1,1).value #账户名
	bank_type = file_name.replace('.xls','')[2:]
	l_zhongxin = [bank_type,count] #银行名和账号放进一个临时list，后面把所有信息一起作为一个list，append进all_info，后面取值方便
	for i in range(sh.nrows):
		if i >= 4:
			if sh.row_values(i)[0] != '':			
				date = sh.row_values(i)[0]	#日期
				
				money_out = sh.row_values(i)[6].replace(',','') #出
				if len(money_out): #长度大于0时说明有值，转为float类型
					money_out = float(money_out)

				money_in = sh.row_values(i)[7].replace(',','') #入
				if len(money_in):
					money_in = float(money_in)

				money_now = sh.row_values(i)[8].replace(',','') #结余
				if len(money_now):
					money_now = float(money_now)

				to_count_name = sh.row_values(i)[4] #对方账号名
				to_count = sh.row_values(i)[3] #对方账户
				beizhu = sh.row_values(i)[2] #备注
	
				l_zhongxin.append({"count":count,"date":date,"money_out":money_out,"money_in":money_in,"money_now":money_now,"to_count_name":to_count_name,"to_count":to_count,"beizhu":beizhu})
				
				# l.append((count,date,money_out,money_in,money_now,to_count_name,to_count,beizhu))
	global all_infos
	all_infos.append(l_zhongxin)
	print('读取',file_name,'成功，准备合并数据......')

#写excel，
def write_excel(all_infos):
	print('开始合并数据......')
	wbk = xlwt.Workbook(encoding='utf-8')
	
	sheet = wbk.add_sheet('sheet 1')
	font = xlwt.Font() # Create Font
	font.bold = True # 加粗

	

	style_title = xlwt.XFStyle() # 标题加粗居中宋体
	borders= xlwt.Borders()
	borders.left = 1
	borders.right = 1
	borders.top = 1
	borders.bottom = 1
	style_title.borders = borders
	alignment = xlwt.Alignment()
	alignment.horz = xlwt.Alignment.HORZ_CENTER
	alignment.vert = xlwt.Alignment.VERT_CENTER
	style_title.alignment = alignment
	font = xlwt.Font()
	font.name = '宋体'
	font.bold = True
	style_title.font = font # font属性添加进style，否则字体设置无效

	style_content = xlwt.XFStyle() # 正文居中宋体
	alignment = xlwt.Alignment()
	alignment.horz = xlwt.Alignment.HORZ_CENTER
	alignment.vert = xlwt.Alignment.VERT_CENTER
	style_content.alignment = alignment
	font = xlwt.Font()
	font.name = '宋体'
	style_content.font = font

	style_num_align = xlwt.XFStyle() # 数字居中Times New Roman
	alignment = xlwt.Alignment()
	alignment.horz = xlwt.Alignment.HORZ_CENTER
	alignment.vert = xlwt.Alignment.VERT_CENTER
	style_num_align.alignment = alignment
	font = xlwt.Font()
	font.name = 'Times New Roman'
	style_num_align.font = font

	style_num = xlwt.XFStyle() # 数字不居中Times New Roman
	alignment = xlwt.Alignment()
	alignment.horz = xlwt.Alignment.HORZ_CENTER
	alignment.vert = xlwt.Alignment.VERT_CENTER
	style_num.alignment = alignment
	font = xlwt.Font()
	font.name = 'Times New Roman'
	style_num.font = font
	
	sheet.write_merge(0, 0, 0, 10, '现金银行存款日记账',style_title)
	sheet.write_merge(1, 2, 0, 2, '2017年度',style_title)#合并行行列列
	sheet.write_merge(3, 4, 0, 0, '月',style_title)
	sheet.write_merge(3, 4, 1, 1, '日',style_title)
	sheet.write_merge(3, 4, 2, 2, '编号',style_title)
	sheet.write_merge(1, 4, 3, 3, '摘要',style_title)
	sheet.write_merge(1, 4, 4, 4, '银行名称',style_title)
	sheet.write_merge(1, 4, 5, 5, '类型',style_title)
	sheet.write_merge(1, 3, 6, 8, '合计',style_title)
	sheet.write_merge(1, 3, 9, 11, '库存现金',style_title)

	sheet.write(4, 6, '收',style_title)
	sheet.write(4, 7, '付',style_title)
	sheet.write(4, 8, '结存',style_title)
	sheet.write(4, 9, '收',style_title)
	sheet.write(4, 10, '付',style_title)
	sheet.write(4, 11, '结存',style_title)
	sheet.write_merge(1, 1, 12, 11+3*len(all_infos), '银行存款',style_title)#
	info_count = 0 #初始化一个变量来保存已经写入的行数，因为不同银行所在列不同
	all_money_ins = []
	all_money_outs = []
	all_money_nows = []
	for i in range(len(all_infos)):
		sheet.write_merge(2, 2, 12+3*i, 14+3*i, all_infos[i][0][0:4],style_title)
		sheet.write_merge(3, 3, 12+3*i, 14+3*i, all_infos[i][1],style_title)
		sheet.write(4, 12+3*i, '收',style_title)
		sheet.write(4, 13+3*i, '付',style_title)
		sheet.write(4, 14+3*i, '结存',style_title)
		# all_money_ins = all_money_ins + all_infos[i][-1]['money_in'] if all_infos[i][-1]['money_in'] != '' else 0
		# all_money_outs = all_money_ins + all_infos[i][-1]['money_out'] if all_infos[i][-1]['money_in'] != '' else 0
		# all_money_nows = all_money_ins + all_infos[i][-1]['money_now'] if all_infos[i][-1]['money_in'] != '' else 0

		all_money_in = Decimal('0.00')
		all_money_out = Decimal('0.00')

		for j in range(len(all_infos[i]) - 2):
			#下面需要通过正则匹配获取日期中正整数，如20170102，取出01时要输出1
			sheet.write(5+j + info_count, 0, re.findall(r"[1-9]\d*",str(all_infos[i][j+2]['date'])[4:6]),style_num_align) 
			sheet.write(5+j + info_count, 1, re.findall(r"[1-9]\d*",str(all_infos[i][j+2]['date'])[6:8]),style_num_align)
			sheet.write(5+j + info_count, 2, j + info_count + 1,style_num_align)
			sheet.write(5+j + info_count, 3, all_infos[i][j+2]['beizhu'],style_content)
			sheet.write(5+j + info_count, 4, all_infos[i][0][0:4],style_content)
			sheet.write(5+j + info_count, 5, all_infos[i][0][4:],style_content)

			#第一个参数，行，j为其中一个银行流水线数据中第j行，第一个银行数据输入完后，要把行数存到info_count，下一个银行要从下面继续输入，
			#因为不同银行的同一类型数据要写在不同行不同列，第二个参数也要自动确认列数
			sheet.write(5+j + info_count, 12+3*i, all_infos[i][j+2]['money_in'],style_num)
			sheet.write(5+j + info_count, 13+3*i, all_infos[i][j+2]['money_out'],style_num)
			sheet.write(5+j + info_count, 14+3*i, all_infos[i][j+2]['money_now'],style_num)
			sheet.write(5+j + info_count, 6, all_infos[i][j+2]['money_in'],style_num)
			sheet.write(5+j + info_count, 7, all_infos[i][j+2]['money_out'],style_num)
			sheet.write(5+j + info_count, 8, all_infos[i][j+2]['money_now'],style_num)
			all_money_in = all_money_in + Decimal(all_infos[i][j+2]['money_in'] if all_infos[i][j+2]['money_in'] !='' else 0)
			all_money_out = all_money_out + Decimal(all_infos[i][j+2]['money_out'] if all_infos[i][j+2]['money_out'] !='' else 0)
			all_money_now = Decimal(all_infos[i][j+2]['money_now'] if all_infos[i][j+2]['money_now'] !='' else 0)
		all_money_ins.append({'all_money_in':all_money_in,'all_money_out':all_money_out,'all_money_now':all_money_now})		
		info_count = info_count + len(all_infos[i]) - 2

	money_ins = Decimal('0.00')
	money_outs = Decimal('0.00')
	money_nows = Decimal('0.00')
	for k in range(len(all_money_ins)):
		money_ins = money_ins + Decimal((all_money_ins[k]['all_money_in']))
		money_outs = money_outs + Decimal((all_money_ins[k]['all_money_out']))
		money_nows = money_nows + Decimal((all_money_ins[k]['all_money_now']))
		sheet.write(5 + info_count, 12 + k * 3, float(all_money_ins[k]['all_money_in']),style_num)
		sheet.write(5 + info_count, 13 + k * 3, float(all_money_ins[k]['all_money_out']),style_num)
		sheet.write(5 + info_count, 14 + k * 3, float(all_money_ins[k]['all_money_now']),style_num)
	sheet.write(5 + info_count, 3, '合计',style_title)
	sheet.write(5 + info_count, 6, float(money_ins),style_num)
	sheet.write(5 + info_count, 7, float(money_outs),style_num)
	sheet.write(5 + info_count, 8, float(money_nows),style_num)
	name = '日记账' + time.strftime('%Y%m%d%H%M%S',time.localtime()) + '.xls'
	wbk.save(name) #循环输入后保存，此时的文件名对应的文件可以存在，会被覆盖，但是不能是打开状态，会报错
	print('数据合并成功！合并后的文件为：',name)

print('使用方法：将各银行原始流水表和本程序放在同一文件夹下，并将流水表按照想要的顺序重命名，如：1-建设银行基本户，数字为顺序，银行名必须是全名，如建设银行，不能写成建行；后面基本户为账户类型，有则写，没有就不写！')
input('按任意键开始程序！')

for i in os.listdir():
	if 'xls' in i:
		if '浦发' in i:
			read_excel_pufa(i)
		if '建设银行' in i or '建行' in i:
			read_excel_jianhang(i)
		if '中信' in i:
			read_excel_zhongxin(i)
		if '招行' in i or '招商银行' in i:
			read_excel_zhaohang(i) 
			
write_excel(all_infos)

input('按任意键退出！')
