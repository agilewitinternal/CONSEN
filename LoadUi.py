from tkinter import *
from mailmerge import MailMerge
import xlrd
import sys
import os
import codecs
from tkinter import filedialog
from tkinter.filedialog import askdirectory

current_row = 0
sheet_num = 0

data_list = []
window = Tk()
window.geometry('650x450')

def exit_label():
	btn = Button(window,text="Exit",width=6,command=sys_exit).grid(row=14,column=0,sticky=E)

def Exl():
	entered_text=exlentry.get()
	filename = filedialog.askopenfilename()
	exlentry.insert(0,str(filename))
	exit_label()
	#btn = Button(window,text="Exit",width=6,command=sys_exit).grid(row=14,column=0,sticky=E)
	data_list.append(filename)
	print(data_list)
def wrd():
	entered_text=wordentry.get()
	filename = filedialog.askopenfilename()
	wordentry.insert(0,str(filename))
	exit_label()
	#btn = Button(window,text="Exit",width=6,command=sys_exit).grid(row=14,column=0,sticky=E)
	data_list.append(filename)
	print(data_list)
def load_dir():
	exit_label()
	#btn = Button(window,text="Exit",width=6,command=sys_exit).grid(row=14,column=0,sticky=E)
	path=askdirectory()
	direntry.insert(0,str(path))
	data_list.append(path)
	print(data_list)
def load_dir_PDF():
	#btn = Button(window,text="Exit",width=6,command=sys_exit).grid(row=14,column=0,sticky=E)
	exit_label()
	path=askdirectory()
	direntrypdf.insert(0,str(path))
	data_list.append(path)
	print(data_list)
def num():
	exit_label()
	#btn = Button(window,text="Exit",width=6,command=sys_exit).grid(row=14,column=0,sticky=E)
	entered_text=numentry.get()
	numentry.insert(0,'entered sheet num: ')
	numentry.insert(0,str(entered_text))
	data_list.append(int(entered_text))
	print(data_list)
def sys_exit():
	window.destroy()
window.title("Variable pass")
window.configure(background='blue')
lbl1 = Label(window,text="Provide a link of Word ")
lbl1.grid(row=0,column=0,sticky=W)
exlentry = Entry(window, width=100,bg='white',textvariable=StringVar())
exlentry.grid(row=1,column=0,sticky=W)
exit_label()
#btn = Button(window,text="Exit",width=6,command=sys_exit).grid(row=14,column=0,sticky=E)
btn = Button(window,text="Submit",width=6,command=Exl).grid(row=2,column=0,sticky=W)
lbl2 = Label(window,text="Provide a link of Excel")
lbl2.grid(row=3,column=0,sticky=W)
wordentry = Entry(window, width=100,bg='white')
wordentry.grid(row=4,column=0,sticky=W)
exit_label()
#btn = Button(window,text="Exit",width=6,command=sys_exit).grid(row=14,column=0,sticky=E)
btn = Button(window,text="Submit",width=6,command=wrd).grid(row=5,column=0,sticky=W)
lbl3 = Label(window,text="Provide a directory link  to save the documents")
lbl3.grid(row=6,column=0,sticky=W)
direntry = Entry(window, width=100,bg='white')
direntry.grid(row=7,column=0,sticky=W)
exit_label()
#btn = Button(window,text="Exit",width=6,command=sys_exit).grid(row=14,column=0,sticky=E)
btn = Button(window,text="Submit",width=6,command=load_dir).grid(row=8,column=0,sticky=W)
lbl5 = Label(window,text="Provide a directory link  to save the PDF")
lbl5.grid(row=9,column=0,sticky=W)
direntrypdf = Entry(window, width=100,bg='white')
direntrypdf.grid(row=10,column=0,sticky=W)
exit_label()
#btn = Button(window,text="Exit",width=6,command=sys_exit).grid(row=14,column=0,sticky=E)
btn = Button(window,text="Submit",width=6,command=load_dir_PDF).grid(row=11,column=0,sticky=W)
lbl4 = Label(window,text="Provide a sheet number of the excel")
lbl4.grid(row=12,column=0,sticky=W)
numentry = Entry(window, width=20,bg='white')
numentry.grid(row=13,column=0,sticky=W)
exit_label()
#btn = Button(window,text="Exit",width=6,command=sys_exit).grid(row=14,column=0,sticky=E)
btn = Button(window,text="Submit",width=6,command=num).grid(row=14,column=0,sticky=W)
window.mainloop()
