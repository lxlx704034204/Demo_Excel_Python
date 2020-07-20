#!/usr/bin/env python
	#tab 制表符
    #空格

#from tkinter import tk
import tkinter as tk  # 使用Tkinter前需要先导入
from time import sleep
#from tkinter import *
#from tkMessageBox import showwarning
from tkinter.messagebox import showwarning
import win32com.client as win32

alert = lambda app: showwarning(app, 'Exit?')
RANGE = range(3, 8)

def excel():
	app = 'Excel-ha1' # alert弹窗的title
	#excel = win32.gencache.EnsureDispatch('%s.Application' % app )
	excel = win32.gencache.EnsureDispatch('Excel.Application')
	wb = excel.Workbooks.Add()
	sh = wb.ActiveSheet
	
	"""
	#0  代表隐藏对象，但可以通过菜单再显示
	#-1 代表显示对象
	#2  代表隐藏对象，但不可以通过菜单显示，只能通过VBA修改为显示状态
	"""
	excel.Visible= True
	sleep(0.1)

	sh.Cells(1, 1).Value ='Python-to-%s Demo' % app # A1赋值为1
	sh.Cells(1,1).Font.Bold = True #加粗
	sh.Cells(1, 1).Name = "Arial" # 字体类型
	sh.Range(sh.Cells(1, 1),sh.Cells(1,2)).Font.Name = "Times New Roman" #选择指定区域
	sh.Range(sh.Cells(1, 1), sh.Cells(1,2)).Font.Size = 10.5
	#sh.Rows(row).Delete()#删除行  
	#sh.Columns(row).Delete()#删除列
	sh.Range(sh.Cells(1,1), sh.Cells(1,1)).HorizontalAlignment = win32.constants.xlCenter #水平居中xlCenter
	sleep(0.1)
    
	for i in RANGE:
		sh.Cells(i, 1).Value = 'Line %d' % i
		sleep(0.1)
	sh.Cells(i+2, 1).Value = "Th-th-th-that's all folks!"

	#sh.SaveAs(path+'demo.xls')
	alert(app)
	wb.Close(False)
	excel.Application.Quit()
	
if  __name__ == '__main__':
	window = tk.Tk()
	window.withdraw()
	excel()
