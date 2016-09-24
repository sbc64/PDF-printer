import os, sys
import win32print
from pyexcel_xls import get_data
from openpyxl import load_workbook
import subprocess
import easygui
from tkinter import filedialog, Tk, Label, Button, LEFT, RIGHT, W, Message, StringVar, Toplevel, Listbox, ttk, Text
from tkinter import *
import tkinter
from tkinter import ttk
from shutil import copyfile
import threading
from operator import itemgetter
import win32api

def checkDependencies():
	pass

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

					excelTempDict['DRAWING']=str(row[4]).rstrip()
					excelTempDict['REV']=str(row[5]).rstrip()
					excelTempDict['GAGE']=str(row[8]).rstrip()
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
					excelTempDict['REV']=str(row[5].value).rstrip()
					excelTempDict['GAGE']=str(row[8].value).rstrip()
					dictList.append(excelTempDict)

		else:
			easygui.msgbox("File type not supported")
			return
	return dictList

def findPDFs(completelyOrderedDictList):

	RELEASED_PDFS_LS = os.listdir("T:\RELEASED_FILES\CURRENT_PDF")
	#RELEASED_PDFS_LS = os.listdir("T:\ENGINEERING\FILE_SANDBOX\CADKEY TO PDF")

	foundParts, unfoundParts, wrongRevsion = [], [], []
	for dictx in completelyOrderedDictList:
		for x in range(len(RELEASED_PDFS_LS)):

			if dictx['PARTNO'] in RELEASED_PDFS_LS[x]:
				if ('R'+dictx['REV'] in RELEASED_PDFS_LS[x]):
					foundParts.append("\"T:\RELEASED_FILES\CURRENT_PDF\\"+RELEASED_PDFS_LS[x]+"\"")
					break
				else:
					wrongRevsion.append(dictx['PARTNO'])
					break
			else:
				if(len(RELEASED_PDFS_LS)-1 == x):
					unfoundParts.append(dictx['PARTNO'])

	return [foundParts,unfoundParts, wrongRevsion]

def configurePrinter():
	pass

def sortPartNumberList(excelInfo):

	floatTempList, otherTempList = [], []
	dictsSortedByGage = sorted(excelInfo, key=itemgetter('GAGE'))
	previousGage = None
	gageSeperatedList = []
	allList = []

	for dictx in dictsSortedByGage:
		#print (dictx['GAGE'])

		if previousGage == None:
			previousGage = dictx['GAGE']

		if previousGage == dictx['GAGE']:
			gageSeperatedList.append(dictx)

		else:
			allList.append(gageSeperatedList)
			previousGage = dictx['GAGE']
			gageSeperatedList = []
			gageSeperatedList.append(dictx)

	#captures the last list that did not get appende in if previousGage == dictx['GAGE']:
	allList.append(gageSeperatedList)
	orderedList = []
	for listx in allList:
		#Sorts the lists that contain the same gage by the drawing number
		listSortedByDrawingNumber = sorted(listx, key=itemgetter('DRAWING'))
		for x in listSortedByDrawingNumber:
			orderedList.append(x)
	return orderedList

def ghostscript(pdfPath, jobCounter, printer,paperType):
	"""
	command = r"gswin64c.exe -dPrinted -q -sDEVICE=mswinpr2 -dNoCancel -sPAPERSIZE="+paperType+" -dBATCH -dFitPage -dNOPROMPT -dFIXEDMEDIA -dNOPAUSE -sOutputFile="
	#the following string shoudl generate something like this: -sOutputFile="\\spool\KONICA MINOLTA 423"
	command = command + "\"\\\\spool\\"+printer+"\" "
	command = command + pdfPath
	"""
	command = r"gswin64c.exe -dPrinted -q -dNoCancel -sPAPERSIZE="+paperType+" -dBATCH -dFitPage -dNOPROMPT -dFIXEDMEDIA -dNOPAUSE "
	command = command + pdfPath

	#print (command)
	with open('tmp', 'a') as g:
		gs = subprocess.run(command, shell=True, stdout=g,stderr=g,stdin=None)
	if gs.returncode != None:
		#print ("Finished")
		jobCounter = jobCounter + 1
		return jobCounter

class mainUIClass:

	def __init__(self,master):
		self.master = master
		self.master.title("PDF printer")

		#self.t = Toplevel

		self.printer = 'KONICA MINOLTA 423'
		"""
		sizex = 600
		sizey = 400
		posx  = 500
		posy  = 400
		master.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, posx, posy))
		"""

		self.top_label = Label (master, text="Please select the file you want to print the part numbers from.")
		self.top_label.grid(row=0, column=1, sticky=W)


		self.entryVar = StringVar()
		#self.entryVar.set("Browse for burn file")
		self.path_show = Label(master, width=100, background="white", anchor=W, relief=GROOVE, height=1, textvariable=self.entryVar)
		self.path_show.grid(row=1, column=1,columnspan=2)

		self.browse_button = Button(master, text="Browse", command=self.askFilename)
		self.browse_button.grid(row=1, column=3, sticky=W)


		#Headers
		self.label_found_files_header = Label(master,text="Files that were found:")
		self.label_found_files_header.grid(row=2, column=1,sticky=W)

		self.label_unfound_files_header = Label(master,text="Parts that were NOT found:")
		self.label_unfound_files_header.grid(row=2, column=2,sticky=W)

		self.label_wrong_revision_header = Label(master,text="Parts with wrong revision:")
		self.label_wrong_revision_header.grid(row=5, column=2,sticky=W)



		#Bodys
		self.text_found_files_body = Text(master, background="white",height=40, width=60, relief=GROOVE)
		self.text_found_files_body.grid(row=4, column=1, sticky=E, rowspan=4, padx=3)

		self.text_unfound_files_body = Text(master, background="white",height=19, width=40, relief=GROOVE)
		self.text_unfound_files_body.grid(row=4, column=2, sticky=E, padx=3,columnspan=2)

		self.text_revision_body = Text(master, background="white",height=19, width=40, relief=GROOVE)
		self.text_revision_body.grid(row=6, column=2, sticky=E,padx=3,columnspan=2)

		self.text_found_files_body.configure(state=DISABLED)
		self.text_unfound_files_body.configure(state=DISABLED)
		self.text_revision_body.configure(state=DISABLED)


		#Progress bar


		#Buttons
		self.close_button = Button(master, text="Close", command=master.quit)
		self.close_button.grid(row=15, column=1,sticky=W,pady=5)

		self.options_button = Button(master, text="Options")
		self.options_button.grid(row=15,column=1,pady=5)

		printingThread = threading.Thread(target=self.printFiles)
		self.excecute_button = Button(master, text="Print", command=printingThread.start)
		self.excecute_button.grid(row=15,column=3,sticky=W,pady=5)

	def askFilename(self):
		currdir = os.getcwd()
		filey = None
		filey = filedialog.askopenfilename(parent=self.master, initialdir=currdir, title='Please select burn file')

		if type(filey) == str:
			self.entryVar.set(filey)
			self.entryVar.set(filey)

		excelInformation = readExcel(str(self.entryVar.get()))
		orderedExcelInfo = sortPartNumberList(excelInformation)
		temp = findPDFs(orderedExcelInfo)
		self.PDFs = temp [0]
		self.unfoundItems = temp[1]
		self.wrongRevsion = temp[2]
		#print (temp[2])
		self.text_found_files_body.configure(state=NORMAL)
		self.text_unfound_files_body.configure(state=NORMAL)
		self.text_revision_body.configure(state=NORMAL)


		counter = 0
		for pdf in self.PDFs:
			self.text_found_files_body.insert(str(counter)+'.0',pdf+"\n")
			counter += 1
		counter = 0

		for part in self.unfoundItems:
			self.text_unfound_files_body.insert(str(counter)+'.0', part+"\n")
			counter += 1

		counter = 0
		for part in self.wrongRevsion:
			print (part)
			self.text_revision_body.insert(str(counter)+'.0', part+"\n")
			counter += 1



		self.text_found_files_body.configure(state=DISABLED)
		self.text_unfound_files_body.configure(state=DISABLED)
		self.text_revision_body.configure(state=DISABLED)


	def Alarm(self):
		self.t.focus_force()
		self.t.bell()


	def printersNameList(self):
		listx = []
		for printerDetails in win32print.EnumPrinters(2):
			listx.append(printerDetails[2])
		self.printerList = listx

	def create_options_window(self):
		"""
		self.t.(self.t)
		self.t.wm_title(self.master,"Options")
		self.t.grab_set(self.master)
		self.t.close_button = Button(self.t, text="Close", command=self.master.quit)
		self.t.close_button.pack()
		sizex = 600
		sizey = 400
		posx  = 500
		posy  = 400

		self.t.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, posx, posy))
		self.t.bind("<FocusOut>", self.Alarm)
		"""
		pass

	def printFiles(self):
		for pdf in self.PDFs:
			self.jobCounter = ghostscript(pdf,self.jobCounter,self.printer, 'letter')
			#print (jobCounter)


if __name__=="__main__":

	f = open('tmp','w+')
	f.close()
	root = tkinter.Tk()
	root.style = ttk.Style()
	mainUI = mainUIClass(root)
	root.style.theme_use("winnative")
	root.mainloop()
