# _*_ coding:utf-8 _*_
import xlrd
import xlwt
import xlutils
import os
import re
import time

#定义两个全局list，all_info存储每个excel内容，bank_info存储表格内银行名，以便后面根据银行名来确认写excel时输出位置
all_infos = []
bank_info = []

#读取浦发银行流水线，并把相关信息append进入前面定义的两个list
def read_excel_pufa(file_name):
	wb = xlrd.open_workbook(file_name)
	sh = wb.sheet_by_index(0)
	count = sh.cell(0,1).value #账号
	count_name = sh.cell(1,1).value #账户名
	l_pufa = ['浦发银行' + file_name.replace('.xls',''),count] #银行名和账号放进一个临时list，后面把所有信息一起作为一个list，append进all_info，后面取值方便
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

#读取建设银行流水线，并把相关信息append进入前面定义的两个list
def read_excel_jianhang(file_name):
	
	wb = xlrd.open_workbook(file_name)
	sh = wb.sheet_by_index(0)
	bank_name = sh.cell(3,1).value
	count = sh.cell(4,1).value
	count_name = sh.cell(5,1).value
	l_jianhang = ['建设银行' + file_name.replace('.xls',''),count]
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

#读取中信银行流水线，并把相关信息append进入前面定义的两个list
def read_excel_zhongxin(file_name):
	wb = xlrd.open_workbook(file_name)
	sh = wb.sheet_by_index(0)
	count = sh.cell(1,3).value #账号
	count_name = sh.cell(1,1).value #账户名
	l_zhongxin = ['中信银行' + file_name.replace('.xls',''),count] #银行名和账号放进一个临时list，后面把所有信息一起作为一个list，append进all_info，后面取值方便
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

#写excel，
def write_excel(all_infos):
	wbk = xlwt.Workbook(encoding='utf-8')
	
	sheet = wbk.add_sheet('sheet 1')
	font = xlwt.Font() # Create Font
	font.bold = True # 加粗

	

	style_title = xlwt.XFStyle() # 标题加粗居中宋体
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
	sheet.write_merge(1, 3, 3, 3, '摘要',style_title)
	sheet.write_merge(1, 3, 4, 4, '银行性质（公私）',style_title)
	sheet.write_merge(1, 3, 5, 5, '银行名称',style_title)
	sheet.write_merge(1, 3, 6, 6, '性质（收付）',style_title)
	sheet.write_merge(1, 3, 7, 7, '部门',style_title)
	sheet.write_merge(1, 3, 8, 10, '合计',style_title)
	sheet.write(4, 8, '收',style_title)
	sheet.write(4, 9, '付',style_title)
	sheet.write(4, 10, '结存',style_title)
	sheet.write_merge(1, 1, 14, 13+3*len(all_infos), '银行存款',style_title)#
	info_count = 0 #初始化一个变量来保存已经写入的行数，因为不同银行所在列不同
	
	for i in range(len(all_infos)):
		sheet.write_merge(2, 2, 14+3*i, 16+3*i, all_infos[i][0][6:],style_title)
		sheet.write_merge(3, 3, 14+3*i, 16+3*i, all_infos[i][1],style_title)
		sheet.write(4, 14+3*i, '收',style_title)
		sheet.write(4, 15+3*i, '付',style_title)
		sheet.write(4, 16+3*i, '结存',style_title)
		for j in range(len(all_infos[i]) - 2):
			#下面需要通过正则匹配获取日期中正整数，如20170102，取出01时要输出1
			sheet.write(5+j + info_count, 0, re.findall(r"[1-9]\d*",all_infos[i][j+2]['date'][4:6]),style_num_align) 
			sheet.write(5+j + info_count, 1, re.findall(r"[1-9]\d*",all_infos[i][j+2]['date'][-2:]),style_num_align)
			sheet.write(5+j + info_count, 3, all_infos[i][j+2]['beizhu'],style_content)
			sheet.write(5+j + info_count, 4, '公账',style_content)
			sheet.write(5+j + info_count, 5, all_infos[i][0][0:4],style_content)

			#第一个参数，行，j为其中一个银行流水线数据中第j行，第一个银行数据输入完后，要把行数存到info_count，下一个银行要从下面继续输入，
			#因为不同银行的同一类型数据要写在不同行不同列，第二个参数也要自动确认列数
			sheet.write(5+j + info_count, 14+3*i, all_infos[i][j+2]['money_in'],style_num)
			sheet.write(5+j + info_count, 15+3*i, all_infos[i][j+2]['money_out'],style_num)
			sheet.write(5+j + info_count, 16+3*i, all_infos[i][j+2]['money_now'],style_num)
			sheet.write(5+j + info_count, 8, all_infos[i][j+2]['money_in'],style_num)
			sheet.write(5+j + info_count, 9, all_infos[i][j+2]['money_out'],style_num)
			sheet.write(5+j + info_count, 10, all_infos[i][j+2]['money_now'],style_num)
		info_count = info_count + len(all_infos[i]) - 2

	wbk.save('日记账' + time.strftime('%Y%m%d%H%M%S',time.localtime()) + '.xls') #循环输入后保存，此时的文件名对应的文件可以存在，会被覆盖，但是不能是打开状态，会报错

for i in os.listdir():
	if 'xls' in i:
		if '浦发' in i:
			read_excel_pufa(i)
		if '建设银行' in i or '建行' in i:
			read_excel_jianhang(i)
		if '中信' in i:
			read_excel_zhongxin(i)
			
write_excel(all_infos)

input('计算成功，按任意键退出！')
