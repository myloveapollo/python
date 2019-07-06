#-*-coding:GBK -*-
#制作讲义室随材发放条
import numpy as np 
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill,Border,Side,Alignment,Protection,Font
pd.options.mode.chained_assignment = None


def cell_style(ws,len_index):
	width_dict = {'B':30,'C':6,'D':14.56,'E':12.00,'F':12.00,'G':6,'H':6}
	font = Font(name='宋体',size=20,bold=True)
	thin = Side(border_style='thin',color='00000000')
	alignment = Alignment(horizontal='center',vertical='center',wrap_text=True)
	for row in ws.iter_rows(min_row=1,max_row=len_index,min_col=2,max_col=8):
		for cell in row:
			cell.font=font
			cell.border = Border(bottom=thin)
			cell.alignment = alignment
	for k,v in width_dict.items():
		ws.column_dimensions[k].width = v
	for i in range(len_index+1):
		ws.row_dimensions[i].height = 165
	
def wash_data(filename):
	wb = Workbook()
	data = pd.read_excel(filename, sheet_name='Sheet0',usecols=[4,6,7,9,11,14,17,18,22,25,29])#读取表
	data['已上课次'] = data['已上课次']+1

	weekdays_all = ['周五','周六上午','周六中午','周六下午','周六晚上'
					,'周日上午','周日中午','周日下午','周日晚上','周二','周三','周四']
					
	weekdays_all2 = ['上午','中午','下午','晚上']	
	
	finish_excel = data.loc[2,'教学点']+ '__'+ data.loc[2,'学期']
	
	data_fudao = data.辅导老师.str.replace('[0-9]\d*$','')#删掉辅导老师名称后的数字
	data_fudao2 = data_fudao.str.replace('.*[\u4e00-\u9fa5]','辅导',regex=True)#删掉数字后只剩名字，名字全部替换成辅导
	data.辅导老师 = data_fudao2 +'\n'+ data_fudao #辅导+老师名字
	
	data.教师 =  data.教师.str.replace('[0-9]\d*$','')#删掉老师名称后的数字
	data.教师 = data.教师 + '\n'+data.辅导老师
	# ~ data.教师 = data['教师'].str.cat(data['辅导老师'],join='left',sep=' ')
	data.年级 = data['年级'].str.cat(data['班次'],join='left')
	data.教室 = data.教室.str.replace('[(].*?[)]|[【].*?[】]','')

	data.rename(columns = {'已上课次':'课次','已缴人数':'人数'},inplace=True)	
	data.sort_values(by='上课时间', axis=0, ascending=True,inplace=True)#用上课时间排序方法
	data = data.drop(columns=['辅导老师'],axis=1)
	data = data.drop(columns=['班次','教学点'],axis=1)
	# ~ data = data.drop(columns=['课次'],axis=1)

	if data.学期.iloc[1] == '春季班' or data.学期.iloc[1] == '秋季班':
		for week in weekdays_all:
			data_num = data.loc[data.上课时间.str.contains(week)]
			if week =='周二'or week =='周三' or week =='周四' or week=='周五':
				data_num.sort_values(by='教室',axis=0,ascending=True,inplace=True)
			ws =wb.create_sheet(week+'('+str(len(data_num.index))+'个课)',-1)
			ws.column_dimensions.group('A',hidden=True)
			for r in dataframe_to_rows(data_num,index=False,header=False):
				ws.append(r)
			cell_style(ws,len(data_num.index))
	else:
		for week in weekdays_all2:
			data_num = data.loc[data.上课时间.str.contains(week)]
			ws = wb.create_sheet(week+'('+str(len(data_num.index))+'个课)',-1)
			ws.column_dimensions.group('A',hidden=True)
			for r in dataframe_to_rows(data_num,index=False,header=False):
				ws.append(r)
			cell_style(ws,len(data_num.index))
	finish_excel = finish_excel+'__'+'随材发放条.xlsx'
	wb.save(finish_excel)
	return str(finish_excel)
	

# ~ filename = '大新暑假4.21.xlsx'
# ~ wash_data(filename)









































