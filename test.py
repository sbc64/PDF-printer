import os, sys
import win32print
from pyexcel_xls import get_data
from openpyxl import load_workbook
import subprocess
import easygui
from tkinter import filedialog, Tk, Label, Button, LEFT, RIGHT, W, Message, StringVar, Toplevel, Listbox, ttk
from tkinter import *
import tkinter
from tkinter import ttk


from shutil import copyfile
import threading
from operator import itemgetter
import win32api




def readExcel(fileToRead):
	partNumberList, drawingNOList, revLevelList, gageList = [],[],[],[]
	excelTempDict = {'PARTNO':'','DRAWING':'','REV':'','GAGE':''}
	dictList = []
	if fileToRead:

		if fileToRead.split(".")[1] == "xls" or fileToRead.split(".")[1] == "XLS":
			ws = get_data(str(fileToRead))['Sheet1']
			for row in ws:
				excelTempDict = {}
				partnumber = row[0].rstrip()
				if not partnumber == 'partno':
					if (partnumber[-1] == 'F'):
						excelTempDict['PARTNO']=partnumber[:-1]
					elif(partnumber[-1] == 'S'):
						excelTempDict['PARTNO']=partnumber
					else:
						excelTempDict['PARTNO']=partnumber

					excelTempDict['DRAWING']=row[4].rstrip()
					excelTempDict['REV']=row[5].rstrip()
					excelTempDict['GAGE']=row[8].rstrip()
					dictList.append(excelTempDict)

		elif fileToRead.split(".")[1] == "xlsx" or fileToRead.split(".")[1] == "XLSX":
			ws = load_workbook(filename=fileToRead)['Sheet1']
			for row in ws.rows:
				excelTempDict = {}
				partnumber = str(row[0].value.rstrip())
				if not partnumber == 'partno':
					if (partnumber[-1] == 'F'):
						excelTempDict['PARTNO']=partnumber[:-1]
					elif(partnumber[-1] == 'S'):
						excelTempDict['PARTNO']=partnumber
					else:
						excelTempDict['PARTNO']=partnumber

					excelTempDict['DRAWING']=str(row[4].value.rstrip())
					excelTempDict['REV']=str(row[5].value)
					excelTempDict['GAGE']=str(row[8].value)
					dictList.append(excelTempDict)

		else:
			easygui.msgbox("File type not supported")
			return


	return dictList

if __name__=="__main__":
	print(readExcel(r"C:\Users\sebastianb\Desktop\Repositories\pdfprinter\Copy of 9-23 BURN.xlsx"))
