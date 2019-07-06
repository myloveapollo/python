#-*-coding:GBK -*- 
#更改2019年7月6日
import wx
import os
import xlrd
import done2
import done3
import done4
import done5
import done6

class SiteLog(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self,None,title='Excel制作工具0626>>>徐浩峰:13554941602',size=(450,480))
		self.Center()
		self.OpenFile = wx.Button(self,label='打开',pos=(305,5),size=(80,25))
		self.OpenFile.Bind(wx.EVT_BUTTON,self.OnOpenFile)
		self.MakeExcel1 = wx.Button(self,label='前台课表目录',pos=(10,40),size=(120,25))
		self.MakeExcel1.Bind(wx.EVT_BUTTON,self.ReadFile)
		self.MakeExcel2 = wx.Button(self,label='制作随材发放条',pos=(140,40),size=(120,25))
		self.MakeExcel2.Bind(wx.EVT_BUTTON,self.ReadFileA)
		self.MakeExcel3 = wx.Button(self,label='随材需求统计表',pos=(270,40),size=(120,25))
		self.MakeExcel3.Bind(wx.EVT_BUTTON,self.ReadFileB)
		self.MakeExcel4 = wx.Button(self,label='A5教室门前课表',pos=(10,75),size=(120,25))
		self.MakeExcel4.Bind(wx.EVT_BUTTON,self.ReadFileC)
		self.MakeExcel5 = wx.Button(self,label='A4教室门前课表',pos=(140,75),size=(120,25))
		self.MakeExcel5.Bind(wx.EVT_BUTTON,self.ReadFileD)
		self.MakeExcel6 = wx.Button(self,label='排班考勤表制作',pos=(270,75),size=(120,25))
		self.MakeExcel6.Bind(wx.EVT_BUTTON,self.ReadFileE)
		
		self.filesFilter = "Excel files(*.xlsx)|*.xlsx" #|All files (*.*)|*.*
		self.fileDialog = wx.FileDialog(self, message ="选择单个文件", wildcard = self.filesFilter, style = wx.FD_OPEN)
		self.FileName = wx.TextCtrl(self, pos=(5,5),size=(290,25),style=wx.TE_READONLY|wx.TE_RICH2)
		self.FileContent = wx.TextCtrl(self,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))


	def OnOpenFile(self, event):
		fileResult = self.fileDialog.ShowModal()
		if fileResult != wx.ID_OK:
			return
		self.FileName.AppendText("%s" % self.fileDialog.GetPath())
		# ~ wx.TextCtrl(self, value=str(self.fileDialog.GetPath()), pos=(5,5), size=(290,25)) 

	def ReadFile(self, event):#前台表
		try:
			filename = self.fileDialog.GetPath()
			names = done2.wash_data(filename)
			message1 = '制作成功！>>>位置：'+str(os.getcwd()) +'\\' +names
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except FileNotFoundError:
			message1 = '没有打开文件！请点击打开！'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except xlrd.biffh.XLRDError:
			message1 = '打开的Excel表不符合"前台表课表"要求'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		
	def ReadFileA(self, event):#制作讲义室随材发放条
		try:
			filename = self.fileDialog.GetPath()
			names = done3.wash_data(filename)
			message2 = '制作成功！>>>位置：'+str(os.getcwd())+'\\' +names
			wx.TextCtrl(self, value=message2,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except FileNotFoundError:
			message1 = '请点击打开文件！'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except xlrd.biffh.XLRDError:
			message1 = '打开的Excel表不符合“制作随材发放条”要求'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		
	def ReadFileB(self, event):#讲义随材需求量统计表
		try:
			filename = self.fileDialog.GetPath()
			names = done4.wash_data(filename)
			message = '制作成功！>>>位置：'+str(os.getcwd())+'\\' +names
			wx.TextCtrl(self, value=message,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except FileNotFoundError:
			message1 = '没有文件被打开'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except xlrd.biffh.XLRDError:
			message1 = '打开的Excel表不符合“随材需求统计表”要求'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
			
	def ReadFileC(self, event):#教室前课表A5
		try:
			filename = self.fileDialog.GetPath()
			sizeA = 'A5'
			names = done5.final_fuc(filename,sizeA)
			message = '制作成功！>>>位置：'+str(os.getcwd())+'\\'+'A5'+names
			wx.TextCtrl(self, value=message,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except FileNotFoundError:
			message1 = '没有文件被打开'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except xlrd.biffh.XLRDError:
			message1 = '打开的Excel表不符合“教室前课表”要求'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))		

	def ReadFileD(self, event):#教室前课表A4
		try:
			filename = self.fileDialog.GetPath()
			sizeA = 'A4'
			names = done5.final_fuc(filename,sizeA)
			message = '制作成功！>>>位置：'+str(os.getcwd())+'\\' +'A4'+names
			wx.TextCtrl(self, value=message,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except FileNotFoundError:
			message1 = '没有文件被打开'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except xlrd.biffh.XLRDError:
			message1 = '打开的Excel表不符合“教室前课表”要求'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		
	def ReadFileE(self, event):#排班考勤表制作
		try:
			filename = self.fileDialog.GetPath()
			names = done6.wash_data(filename)
			message = '制作成功！>>>位置：'+str(os.getcwd())+'\\' +names
			wx.TextCtrl(self, value=message,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except FileNotFoundError:
			message1 = '没有文件被打开'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		except xlrd.biffh.XLRDError:
			message1 = '打开的Excel表不符合要求'
			wx.TextCtrl(self, value=message1,pos=(5,110),size=(430,480),style=(wx.TE_MULTILINE))
		
		
if __name__=='__main__':
	app = wx.App()
	SiteFrame = SiteLog()
	SiteFrame.Show()
	app.MainLoop()
