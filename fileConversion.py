# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import xlrd
import time


root = tk.Tk()
root.resizable(False,False)

canvas1 = tk.Canvas(root,width=450,height=200,bg='light blue',relief='raised')
canvas1.grid(row=0,column=0)

canvas2 = tk.Canvas(root,width=450,height=140,bg='light blue',relief='raised')
canvas2.grid(row=1,column=0)

title1 = tk.Label(root,text='File Conversion Program',bg='light blue')
title1.config(font=('helvetica',14,'bold'))
canvas1.create_window(230,20, window=title1)

def get_time():
	_time = time.ctime()
	return _time

lblTime = tk.Label(root,font=('Times',12,'bold'),bg='#C5DAFC',text=get_time())
lblTime.config(font=('helvetica',11,'bold'))
canvas1.create_window(230,50, window=lblTime)

def get_excel_file():
	global read_file

	file_path = filedialog.askopenfilename()
	read_file = pd.read_excel(file_path)

browseFileButton = tk.Button(root,width=25,height=2,text='Select Excel File to csv',bg='green',fg='white',
	font=('helvetica',12,'bold'),command=get_excel_file)
canvas1.create_window(230,100,window=browseFileButton)

def convertToCsv():
	global read_file

	file_path = filedialog.asksaveasfilename(defaultextension='.csv')

	read_file.to_csv(file_path,index=None,header=True)

saveFileButton = tk.Button(root,width=25,height=2,text='Convert Excel to csv',bg='yellow',fg='black',
	font=('helvetica',12,'bold'),command=convertToCsv)
canvas1.create_window(230,164,window=saveFileButton)

def convertExcelToText():
	global read_file
	myList = []

	# Give the location of the file
	file_path = filedialog.askopenfilename()
	loc = (file_path)
	  
	wb = xlrd.open_workbook(loc)
	sheet = wb.sheet_by_index(0)
	sheet.cell_value(0, 0)

	for i in range(sheet.nrows):
		line = (sheet.row_values(i))
		myList.append(line)
	
	file_pathsave = filedialog.asksaveasfilename(defaultextension='.txt')
	with open(file_pathsave,'w',encoding='utf-8') as txtFile:

		for x in range(len(myList)):
			myLine = ' '.join(myList[x])
			if x+1 < len(myList):
				txtFile.write(f'%s,' % myLine)
			else:
				txtFile.write(f'%s' % myLine)


convertExcelToTextButton = tk.Button(root,width=25,height=2,text='Select and Convert Excel to txt',bg='pale green',fg='black',
	font=('helvetica',12,'bold'),command=convertExcelToText)
canvas2.create_window(230,40,window=convertExcelToTextButton)


def exitApp():
	MesgBox = tk.messagebox.askquestion('Exit Program','Are you sure you want to exit the program ?')
	if MesgBox == 'yes':
		root.destroy()

exitButton = tk.Button(root,text='Exit Program',width=25,height=2,font=('helvetica',12,'bold'),bg='orange',
	fg='#1346DC',command=exitApp)
canvas2.create_window(230,106,window=exitButton)



root.mainloop()