#-*-coding:GBK -*- 
#讲义随材需求量统计表
import numpy as np
import pandas as pd
from openpyxl import Workbook

from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill,Border,Side,Alignment,Protection,Font
pd.options.mode.chained_assignment = None

def cell_style(ws,len_index,names,names2):
	width_dict={'A':44,'B':9,'C':9,'D':14,'E':9}
	thin = Side(border_style='thin',color='00000000')
	alignment = Alignment(horizontal='center',vertical='center',wrap_text=False)
	alignment1 = Alignment(horizontal='center',vertical='center',wrap_text=True)
	font = Font(name='宋体',size=20,bold=True)
	font1 = Font(bold=True)
	font2 = Font(bold=True,color='FFFF0000')

	ws['A1'] =str(names)+'>>'+names2+'教材需求量统计'+'('+str(len_index)+'种班型)'
	ws['A2'] = '校区所有'+names2+'班级类型'
	ws['B2'] = '所有'+names2+'开班总数量'
	ws['C2'] = '截止目前已缴费人数'
	ws['D2'] = names2+'班型全满教材需求量(该科目所有班的限额相加)'
	ws['E2'] = '多余教材(D列减去C列)'
	ws.merge_cells('A1:E1')

	ws['A1'].font = font
	# ~ ws.row_dimensions[2].height = 30
	for row in ws.iter_rows(min_row=1,max_row=len_index+2,min_col=1,max_col=5):
		for cell in row:
			cell.border = Border(top=thin,left=thin,right=thin,bottom=thin)
			cell.alignment = alignment
	for c in ws[2]:
		c.alignment = alignment1
		c.font = font1
	ws.row_dimensions[2].height = 70
	
	for k,v in width_dict.items():
		ws.column_dimensions[k].width = v

	for i in range(3,len_index+3):
		ws['E'+str(i)].value=ws['D'+str(i)].value - ws['C'+str(i)].value
		if int(ws['E'+str(i)].value)<=0:
			ws['E'+str(i)].font = font2
	for i in range(3,len_index+2):
		ws.row_dimensions[i].height = 14


def wash_data(filename):

	data = pd.read_excel(filename, sheet_name='Sheet0',usecols=[4,6,7,9,11,17,18,22,24,29,31])#读取表
	finish_excel = data.loc[2,'教学点']+ '__'+ data.loc[2,'学期']
	data.rename(columns = {'学期':'班级类型'},inplace=True)
	data.班次 = data.班次.str.replace('^[\u4e00-\u9fa5][\u4e00-\u9fa5][\u4e00-\u9fa5]|双师','')
	
	class_dic = {'小学一':'<1>一','小学二':'<2>二','小学三':'<3>三','小学四':'<4>四','小学五':'<5>五','小学六':'<6>六',
				'初中一':'<7>初一','初中二':'<8>初二','初中三':'<9>初三'}
	for k,v in class_dic.items():
		data.年级 = data['年级'].str.replace(k,v)
		
	data.班级类型 = data['班级类型'].str.cat(data['年级'],join='left',sep=' ')
	data.班级类型 = data['班级类型'].str.cat(data['学科'],join='left',sep=' ')
	data.班级类型 = data['班级类型'].str.cat(data['班次'],join='left',sep=' ')
	data.班次 = data.班次.str.replace('^[\u4e00-\u9fa5][\u4e00-\u9fa5][\u4e00-\u9fa5]','')
	
	
	data = data.drop(columns=['年级','班次','教师','教室','教学点','上课时间','总课次'])
	banshu  = data.班级类型.value_counts()
	banshu = pd.DataFrame(banshu)

	tongji = data.groupby('班级类型').sum()

	frames = [banshu,tongji]
	result = pd.concat(frames,axis=1,sort=False)

	
	course = ['数学','英语','语文','物理','化学','科学']
	wb = Workbook()
	for c in course:
		ws = wb.create_sheet(c,-1)
		result1 = result.loc[result.index.str.contains(c)]
		result1 = result1.sort_index()
		for r in dataframe_to_rows(result1,index=True,header=True):
			ws.append(r)
		cell_style(ws,len(result1.index),finish_excel,c)
	finish_excel = finish_excel+'__'+'随材需求统计表.xlsx'
	wb.save(finish_excel)
	return str(finish_excel)


# ~ filename = '大新春季.xlsx'
# ~ wash_data(filename)










































