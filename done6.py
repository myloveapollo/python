#-*-coding:GBK -*- 
#运营考勤排班导入表
import re
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill,Border,Side,Alignment,Protection,Font,numbers
from openpyxl.worksheet.datavalidation import DataValidation
pd.options.mode.chained_assignment = None

def cell_style(ws,len_index):
	width_dict={'A':15.13,'B':15.13,'C':19.28,'D':15.13,'E':15.13,'F':15.13}
	thin = Side(border_style=None,color='00000000')
	alignment = Alignment(horizontal='center',vertical='center',wrap_text=False)
	font = Font(name='宋体',size=11,bold=False)
	
	list_content = ('"06:00,06:30,07:00,07:30,08:00,08:30,09:00,09:30,10:00,10:30,11:00,11:30,'
					'12:00,12:30,13:00,13:30,14:00,14:30,15:00,15:30,16:00,16:30,17:00,17:30,'
					'18:00,18:30,19:00,19:30,20:00,20:30,21:00,21:30,22:00,22:30,23:00,23:30"')#00:00,00:30,01:00,01:30,02:00,02:30,03:00,03:30,04:00,04:30,05:00,05:30,'

	dv_type = DataValidation(type="list", formula1='"年假,工作,休息,入离职缺勤,培训,病假,医疗期,事假,婚假,产假,产检,哺乳假,丧假"', allow_blank=False)
	dv_time = DataValidation(type="list", formula1=list_content,operator='equal', allow_blank=True)
	ws.add_data_validation(dv_type)
	ws.add_data_validation(dv_time)
	type_index = 'D2:D'+str(len_index+1)
	time_index = 'E2:F'+str(len_index+1)
	dv_type.add(type_index)
	dv_time.add(time_index)
	
	for row in ws.iter_rows(min_row=2,max_row=len_index+1,min_col=3,max_col=3):
		for cell in row:
			cell.number_format = 'yyyy-mm-dd'
			
	for row in ws.iter_rows(min_row=2,max_row=len_index+1,min_col=5,max_col=6):
		for cell in row:
			cell.number_format = 'hh:mm'
	
	for row in ws.iter_rows(min_row=1,max_row=len_index+1,min_col=1,max_col=6):
		for cell in row:
			cell.border = Border(top=thin,left=thin,right=thin,bottom=thin)
			cell.alignment = alignment
			cell.font = font

	for k,v in width_dict.items():
		ws.column_dimensions[k].width = v


def ceshi(data):
	writer = pd.ExcelWriter('ceshi.xlsx')
	data.to_excel(writer,index=True,header=True)
	writer.save()


def wash_data(filename):
	names=['工号','姓名','周一','周二','周三','周四','周五','周六','周日']
	data = pd.read_excel(filename,sheet_name=0,header=None,names=names,usecols=[1,2,3,5,7,9,11,13,15])#读取表,

	data_time = data.iloc[1]#提取日期已备用,格式要正常的
	name_a = str(data.iloc[1,2])[6:10]+'至'+str(data.iloc[1,8])[6:10]
	in_excel = name_a + '运营考勤排班导入.xlsx' #命名最后生成的表 
	
	# ~ data.drop(axis=1,columns=[0],inplace=True)
	data.drop(axis=0,index=[1,0,2,3,4],inplace=True)#删除没有的行信息
	data.dropna(axis=0,how='all',inplace=True)#删除全空的行信息
	data.replace('OFF','OFF-OFF',regex=True,inplace=True)#OFF项等于空 正则表达式子，不区分大小写
	data = data.T#倒置

	#起草一个表
	data_make = pd.DataFrame(np.full([7*len(data.columns),6],np.nan),columns=['工号','姓名','日期(YYYY-MM-DD)','类型','上班时间','下班时间'])
	
	data_job_num = []
	for nam in list(data.iloc[0]):
		for i in range(1,8):
			data_job_num.append(nam)
	data_make['工号'] = data_job_num

	#获取姓名
	data_col_name = []
	for nam in list(data.iloc[1]):
		for i in range(1,8):
			data_col_name.append(nam)
	data_make['姓名']= pd.Series(data_col_name)#导入姓名

	#获取时间
	data.drop(axis=0,index='姓名',inplace=True)
	data.drop(axis=0,index='工号',inplace=True)
	data_col_arrivetime = []#上班时间
	data_col_leavetime = []#下班时间
	data_type =[]#类型

	for col in list(data.columns):
		data_split_f = data[col].str.split('-').str[0]#上班时间用前面的
		data_split_s = data[col].str.split('-').str[1]#下班时间用后面的
		for f in list(data_split_f):
			if len(str(f))==4:#len(f) == 4:
				f = '0'+f
			data_col_arrivetime.append(f)
		for s in list(data_split_s):
			data_col_leavetime.append(s)
			if s !='OFF':
				data_type.append('工作')
			else:
				data_type.append('休息')

	data_make['类型'] = data_type#导入类型数据
	data_make['上班时间'] = data_col_arrivetime#导入上班时间
	data_make['下班时间'] = data_col_leavetime#导入下班时间
	data_make.replace('OFF|off','',regex=True,inplace=True)
	
	# ~ #获取'日期(YYYY-MM-DD)'--------方法1
	# ~ data_time = list(map(lambda x: '2019-'+x, data_time))#用公式在每个字符前加2019-
	
	data_yyyy= []
	for data_t in range(0,len(list(data.columns))):
		for data_t in list(data_time[2:]):
			data_yyyy.append(str(data_t)[:-9])#按每个col，datatime次添加
	data_make['日期(YYYY-MM-DD)'] = data_yyyy #赋值给col日期
	
	#用openpyxl设置格式
	wb = Workbook()
	ws = wb.create_sheet('Sheet1',-1)
	for r in dataframe_to_rows(data_make,index=False,header=True):
		ws.append(r)
	cell_style(ws,len(data_make.index))
	wb.save(in_excel)
	return in_excel
	
# ~ filename = '黄锦荣.xlsx'
# ~ wash_data(filename)


