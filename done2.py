#-*-coding:GBK -*- 
#前台课表
import numpy as np#导入nump数据模块
import pandas as pd
import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill,Border,Side,Alignment,Protection,Font
pd.options.mode.chained_assignment = None


def cell_style0(ws0,len_index0,week0,names):
	width_dict={'A':6.7,'B':11,'C':5,'D':20,'E':16,'F':15,'G':11,'H':11,'I':21,'J':7.5,'K':5}
	thin = Side(border_style='thin',color='00000000')
	alignment = Alignment(horizontal='center',vertical='center',wrap_text=False)
	font = Font(name='宋体',size=20,bold=True)
	font1 = Font(bold=True)
	font2 = Font(bold=False,color='FFFF0000')
	
	ws0.insert_rows(1)
	ws0['A1']= names+' '+week0+'课表('+str(len_index0)+'个课)'
	ws0.merge_cells('A1:K1')
	ws0['A1'].font = font
	ws0.row_dimensions[1].height = 30
	
	for row in ws0.iter_rows(min_row=1,max_row=len_index0+2,min_col=1,max_col=11):
		for cell in row:
			if cell.value == datetime.datetime.now().strftime('%Y-%m-%d'):#标红与当前日期相等的结课时间
				cell.font = font2
			cell.border = Border(top=thin,left=thin,right=thin,bottom=thin)
			cell.alignment = alignment
	for c in ws0[2]:
		c.font = font1
	for k,v in width_dict.items():
		ws0.column_dimensions[k].width = v
	for i in range(2,len_index0+3):
		ws0.row_dimensions[i].height = 16



def wash_data(filename):
	wb = Workbook()
	data = pd.read_excel(filename, sheet_name='Sheet0',usecols=[4,6,7,9,11,14,17,18,20,21,22,24,29])#读取表
	
	weekdays_all = ['周五','周六上午','周六中午','周六下午','周六晚上'
					,'周日上午','周日中午','周日下午','周日晚上','周二','周三','周四']
					
	weekdays_all2 = ['上午','中午','下午','晚上']	
	finish_excel = data.loc[2,'教学点']+ '__'+ data.loc[2,'学期']
	data.教师 =  data.教师.str.replace('[0-9]\d*$','')#删掉老师名称后的数字

	data_fudao = data.辅导老师.str.replace('[0-9]\d*$','')#删掉辅导老师名称后的数字
	data_fudao2 = data_fudao.str.replace('.*[\u4e00-\u9fa5]','辅导:',regex=True)#删掉数字后只剩名字，名字全部替换成辅导
	data.辅导老师 = data_fudao2 + data_fudao #辅导+老师名字
	
	
	data.教师 = data.教师.str.cat(data.辅导老师,join='left',sep=' ')#老师与辅导老师合二为一
	data.教室 = data.教室.str.replace('[(].*?[)]|[【].*?[】]','',regex=True)
	# ~ data.班次 = data.班次.str.replace('^[\u4e00-\u9fa5][\u4e00-\u9fa5][\u4e00-\u9fa5]','')
	data.rename(columns = {'已缴人数':'人数'},inplace=True)	#'已上课次':'课次',
	data.sort_values(by='上课时间', axis=0, ascending=True,inplace=True)#用上课时间排序方法
	data = data.drop(columns=['辅导老师','教学点'],axis=1)


	if data.学期.iloc[1] == '春季班' or data.学期.iloc[1] == '秋季班':
		for week in weekdays_all:
			data_num = data.loc[data.上课时间.str.contains(week)]
			if week =='周二'or week =='周三' or week =='周四' or week=='周五':
				data_num.sort_values(by='教室',axis=0,ascending=True,inplace=True)
			else:
				data_num.sort_values(by=['上课时间','教室'],axis=0,ascending=True,inplace=True)
			ws = wb.create_sheet(week+'('+str(len(data_num.index))+'个课)',-1)
			for r in dataframe_to_rows(data_num,index=False,header=True):
				ws.append(r)
			cell_style0(ws,len(data_num.index),week,finish_excel)
			
	elif data.学期.iloc[1] == '短期班' or data.学期.iloc[1] == '活动类' or data.学期.iloc[1] == '诊断类':
		week = data.学期.iloc[1]
		ws = wb.create_sheet(week,-1)
		data.sort_values(by='结课日期',axis=0,ascending=True, inplace=True)#按结课日期排序
		for r in dataframe_to_rows(data,index=False, header=True):
			ws.append(r)
		cell_style0(ws,len(data.index),week,finish_excel)
	
	else:#if data.学期.iloc[1] == '暑假班' or data.学期.iloc[1] == '寒假班'
		data['上课时间'] = data['上课时间'].str.replace('一','1')
		data['上课时间'] = data['上课时间'].str.replace('二','2')
		data['上课时间'] = data['上课时间'].str.replace('三','3')
		data['上课时间'] = data['上课时间'].str.replace('四','4')
		data['上课时间'] = data['上课时间'].str.replace('零','0')
		for week in weekdays_all2:
			data_num = data.loc[data.上课时间.str.contains(week)]
			data_num.sort_values(by=['上课时间','教室'],axis=0,ascending=True,inplace=True)
			ws = wb.create_sheet(week+'('+str(len(data_num.index))+'个课)',-1)
			for r in dataframe_to_rows(data_num,index=False,header=True):
				ws.append(r)
			cell_style0(ws,len(data_num.index),week,finish_excel)
	finish_excel = finish_excel +'__'+'前台课表.xlsx'
	wb.save(finish_excel)
	return str(finish_excel)

# ~ filename = '大新短期班.xlsx'
# ~ wash_data(filename)








































