#!/usr/bin/env python
#encoding=utf-8
import wx
import os
import xlrd
import xlwt

class MyFrame(wx.Frame):
	def __init__(self,parent,title):
		wx.Frame.__init__(self,parent,title=title,size=(200,100))
		panel=wx.Panel(self)
		self.button=wx.Button(panel,-1,u'合并',pos=(50,50))
		self.control = wx.TextCtrl(panel,-1,"",pos=(10,10), size=(180,30))
		self.Bind(wx.EVT_BUTTON, self.data_write, self.button)
		self.Show(True)
	def data_write(self, evt):
		dic = self.control.GetValue()
		print dic
		os.chdir(dic)
		wbk = xlwt.Workbook()
		new_sheet = wbk.add_sheet("sheet1")
		m = 0
		for i in os.listdir(dic):
			i = i.strip("\n")
			table = xlrd.open_workbook(i)
			print i
			sheet = table.sheet_by_index(0)
			for r in range(sheet.nrows):
				for c in range(sheet.ncols):
					new_sheet.write(m,c,sheet.cell(r,c).value)
				m += 1
		wbk.save("111.xls")
		
app = wx.App(False)
frame = MyFrame(None, 'Small editor')
app.MainLoop()
