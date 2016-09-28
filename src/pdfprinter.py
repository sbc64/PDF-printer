import os, sys
import win32print
from xlrd import open_workbook
from openpyxl import load_workbook
import subprocess
from tkinter import filedialog, Tk, Label, Button, LEFT, RIGHT, W, Message, StringVar, Toplevel, Listbox, ttk, Text
from tkinter import *
import tkinter
from tkinter import ttk
import threading
from operator import itemgetter
import io
from time import asctime
#import win32api
import winreg


def checkDependencies():
	pass

def checkGhostScriptPath():
	try:
		pathKey = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Session Manager\Environment"
		key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,"System\CurrentControlSet\Control\Session Manager\Environment",0,winreg.KEY_ALL_ACCESS)
		oldPath = winreg.QueryValueEx(key, "Path")
		dirsList = oldPath[0].rsplit(';')
		ghostscriptPath = r"C:\Program Files\gs\gs9.19"
		pathIsPresent = False
		for x in dirsList:
			if ghostscriptPath == x:
				pathIsPresent = True

		if not pathIsPresent:
			newPath =oldPath[0]+";"+ghostscriptPath
			winreg.SetValueEx(key,"Path",0,2,newPath)

		winreg.CloseKey(key)
	except:
		pass

def readExcel(fileToRead):
	partNumberList, drawingNOList, revLevelList, gageList = [],[],[],[]
	excelTempDict = {'PARTNO':'','DRAWING':'','REV':'','GAGE':''}
	dictList = []
	if fileToRead:
		if fileToRead.split(".")[1] == "xls" or fileToRead.split(".")[1] == "XLS":
			wb = open_workbook(str(fileToRead))
			ws = wb.sheets()[0]
			for row in range(ws.nrows):
				excelTempDict = {}
				partnumber = str(ws.cell(row,0).value).rstrip()
				if not partnumber == 'partno':
					if (partnumber[-1] == 'F'):
						excelTempDict['PARTNO']=partnumber[:-1]
					elif(partnumber[-1] == 'S'):
						excelTempDict['PARTNO']=partnumber
					else:
						excelTempDict['PARTNO']=partnumber

					excelTempDict['DRAWING']=str(ws.cell(row,4).value).rstrip()
					excelTempDict['REV']=str(ws.cell(row,5).value).rstrip()
					excelTempDict['GAGE']=str(ws.cell(row,8).value).rstrip()
					dictList.append(excelTempDict)

		elif fileToRead.split(".")[1] == "xlsx" or fileToRead.split(".")[1] == "XLSX":
			ws = load_workbook(filename=str(fileToRead))['Sheet1']
			for row in ws.rows:
				excelTempDict = {}
				partnumber = str(row[0].value).rstrip()
				if not partnumber == 'partno':
					if (partnumber[-1] == 'F'):
						excelTempDict['PARTNO']=partnumber[:-1]
					elif(partnumber[-1] == 'S'):
						excelTempDict['PARTNO']=partnumber
					else:
						excelTempDict['PARTNO']=partnumber

					excelTempDict['DRAWING']=str(row[4].value).rstrip()
					excelTempDict['REV']=str(row[5].value).rstrip()
					excelTempDict['GAGE']=str(row[8].value).rstrip()
					dictList.append(excelTempDict)
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
		#Sebastian
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


class mainUIClass:

	def __init__(self,master):
		self.master = master
		self.master.title("PDF printer")
		self.selectedPrinter = StringVar()
		self.selectedPaper = StringVar()
		self.selectedPaper.set('letter')
		self.jobCounter = 0
		self.PDFs = None
		self.user = os.getlogin()
		self.POPENFile = 'C:\\Users\\'+self.user+'\\tmpgs'
		#self.t = Toplevel

		self.printer = 'KONICA MINOLTA 423'
		"""
		sizex = 600
		sizey = 400
		posx  = 500
		posy  = 400
		master.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, posx, posy))
		"""

		self.top_label = Label (master, text="Browse for the spreadsheet that contains the partnumbers.")
		self.top_label.grid(row=0, column=1, sticky=W)


		self.entryVar = StringVar()
		#self.entryVar.set("Browse for burn file")
		self.path_show = Label(master, width=100, background="white", anchor=W, relief=GROOVE, height=1, textvariable=self.entryVar)
		self.path_show.grid(row=1, column=1,columnspan=2, padx=5)

		self.browse_button = Button(master, text="Browse", command=self.askFilename)
		self.browse_button.grid(row=1, column=3, sticky=E, padx=10)


		#Headers
		self.label_found_files_header = Label(master,text="Files that were found:")
		self.label_found_files_header.grid(row=3, column=1,sticky=W)

		self.label_unfound_files_header = Label(master,text="Parts that were NOT found:")
		self.label_unfound_files_header.grid(row=3, column=2,sticky=W)

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
		#self.progress = ttk.Progressbar(master, orient="horizontal",length=400, mode="determinate")
		#self.progress.grid(row=14,column=1, columnspan=4, pady=20,sticky=W+E, padx=10)

		#Buttons
		self.close_button = Button(master, text="Close", command=master.quit)
		self.close_button.grid(row=15, column=1,sticky=W,pady=5,padx=4)

		self.options_button = Button(master, text="Printer Settings",command=self.create_options_window)
		self.options_button.grid(row=15,column=1,pady=5,sticky=E,padx=10)


		self.excecute_button = Button(master, text="Print", command=self.checkSettingsBeforePrint)
		self.excecute_button.grid(row=15,column=3,sticky=E,pady=5,padx=10)


	def checkSettingsBeforePrint(self):

		if self.selectedPrinter.get != '' and self.PDFs !=None :
			printingThread = threading.Thread(target=self.printFiles)
			printingThread.start()
		else:
			posx  = 500
			posy  = 400
			sizex = 500
			sizey = 100
			top = Toplevel()
			top.grid_rowconfigure(0,weigh=1)
			top.grid_columnconfigure(0, weight=1)
			top.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, posx, posy))
			top.title("No file loaded")
			msg = Message(top, text="Browse for a file before printing.",width=200, pady=10)
			msg.grid(row=0, column=0,columnspan=5)
			button = Button(top,text="Ok", command=top.destroy)
			button.grid(row=1, column=0)
			return None


	def askFilename(self):
		currdir = os.getcwd()
		filey = None
		filey = filedialog.askopenfilename(parent=self.master, initialdir=currdir, title='Select burn file')

		if type(filey) == str:
			self.entryVar.set(filey)
			try:
				if not (str(self.entryVar.get()).split(".")[1] == "xls" or str(self.entryVar.get()).split(".")[1] == "XLS" or str(self.entryVar.get()).split(".")[1] == "xlsx" or str(self.entryVar.get()).split(".")[1] == "XLSX"):
					top = Toplevel()
					top.title("Wrong file type")
					msg = Message(top, text="You selected a wrong file type.\nPlease use xls or xlsx.", width=300, anchor=CENTER)
					msg.grid(row=0, column=0)
					button = Button(top,text="Ok", command=top.destroy)
					button.grid(row=1, column=0)
					self.entryVar.set("")
					return None
			except:
				top = Toplevel()
				top.title("Wrong file type")
				msg = Message(top, text="You selected a wrong file type.\nPlease use xls or xlsx.", width=300, anchor=CENTER)
				msg.grid(row=0, column=0)
				button = Button(top,text="Ok", command=top.destroy)
				button.grid(row=1, column=0)
				self.entryVar.set("")
				return None
		else:
			return None


		orderedExcelInfo = sortPartNumberList(readExcel(str(self.entryVar.get())))
		temp = findPDFs(orderedExcelInfo)
		self.PDFs = temp [0]
		self.unfoundItems = temp[1]
		self.wrongRevsion = temp[2]

		#Populate the text fields
		self.text_found_files_body.configure(state=NORMAL)
		self.text_unfound_files_body.configure(state=NORMAL)
		self.text_revision_body.configure(state=NORMAL)

		counter = 0
		for pdf in self.PDFs:
			self.text_found_files_body.insert(str(counter)+'.0',pdf.replace("\"","")+"\n")
			counter += 1

		self.totalFiles = counter
		counter = 0

		for part in self.unfoundItems:
			self.text_unfound_files_body.insert(str(counter)+'.0', part.replace("\"","")+"\n")
			counter += 1

		counter = 0
		for part in self.wrongRevsion:
			self.text_revision_body.insert(str(counter)+'.0', part.replace("\"","")+"\n")
			counter += 1

		self.text_found_files_body.configure(state=DISABLED)
		self.text_unfound_files_body.configure(state=DISABLED)
		self.text_revision_body.configure(state=DISABLED)


	def Alarm(self,event):
		self.options_windows.focus_force()
		self.options_windows.bell()

	def create_options_window(self):
		listx = []

		for printerDetails in win32print.EnumPrinters(2):
			listx.append(printerDetails[2])
		self.printerList = listx

		self.options_windows = tkinter.Toplevel(self.master)
		self.options_windows.title("Options")

		counter = 1
		if self.selectedPrinter.get() == '':
			self.selectedPrinter.set(self.printerList[0])

		self.options_label_printer = Label(self.options_windows,text="Select Printer")
		self.options_label_printer.grid(row=0,column=0, sticky=W)

		self.options_label_paper = Label(self.options_windows,text="Select Paper")
		self.options_label_paper.grid(row=0,column=1, sticky=W)

		for printer in self.printerList:
			tempRadio = Radiobutton(self.options_windows, text=printer, variable=self.selectedPrinter, value=printer)
			tempRadio.grid(row=counter, column=0, sticky=W)
			counter += 1

		self.selectedPaper.set('letter')
		self.letter_radio = Radiobutton(self.options_windows, text='8 1/2x11', variable=self.selectedPaper, value='letter')
		self.letter_radio.grid(row=1,column=1,sticky=W)
		self.ledger_radio = Radiobutton(self.options_windows, text='11x17', variable=self.selectedPaper, value='ledger')
		self.ledger_radio.grid(row=2,column=1,sticky=W)


		local_close_button = Button(self.options_windows, text="Ok", command=self.save_options_and_destroy_options_window)
		local_close_button.grid(column=0,row=counter)

		self.options_windows.focus_force()
		self.options_windows.bind("<FocusOut>", self.Alarm)

	def save_options_and_destroy_options_window(self):
		f = open('C:\\Users\\'+self.user+'\\pdf_printer_settings','w+')
		f.write(self.selectedPrinter.get())
		f.close()
		self.options_windows.destroy()


	def printFiles(self):

		for pdf in self.PDFs:
			self.jobCounter = ghostscript(pdf, self.jobCounter, self.selectedPrinter.get(), self.selectedPaper.get(),self.POPENFile)
			#print (jobCounter)

		#os.remove(self.POPENFile)
		posx  = 500
		posy  = 400
		sizex = 500
		sizey = 100
		top = Toplevel()
		top.title("Done")
		top.grid_rowconfigure(0,weigh=1)
		top.grid_columnconfigure(0, weight=1)
		top.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, posx, posy))
		msg = Message(top, text="Sent all files to printer.\nPlease wait for the printer to finish", width=200, pady=10)
		msg.grid(row=0, column=0, columnspan=4)
		button = Button(top,text="Ok", command=top.quit)
		button.grid(row=1, column=0)

def ghostscript(pdfPath, jobCounter, printer,paperType,filex):
	filex = open(filex,'w+')
	filex.write(asctime()+"\n")
	filex.write(str(jobCounter)+"\n")
	filex.write(printer+"\n")
	filex.write(pdfPath+"\n")
	filex.write(paperType+"\n")
	filex.write(str(os.getcwd())+"\n")
	global isFrozen
	if isFrozen:
		if paperType =='letter':
			command = sys._MEIPASS+r"\gswin64c.exe -dPrinted -q -sDEVICE=mswinpr2 -sPAPERSIZE="+paperType+" -dBATCH -dFitPage -dNOPROMPT -dFIXEDMEDIA -dNOPAUSE -sOutputFile="
			#command = r"gswin64c.exe -dPrinted -q -sDEVICE=mswinpr2 -sPAPERSIZE="+paperType+" -dBATCH -dFitPage -dNOPROMPT -dFIXEDMEDIA -dNOPAUSE -sOutputFile="
			#the following string shoudl generate something like this: -sOutputFile="\\spool\KONICA MINOLTA 423"
			command = command + "\"\\\\spool\\"+printer+"\" "
			command = command + pdfPath
		else:
			command = sys._MEIPASS+r"\gswin64c.exe -dPrinted -q -sDEVICE=mswinpr2 -sPAPERSIZE="+paperType+" -dBATCH -dFitPage -dNOPROMPT -dFIXEDMEDIA -dNOPAUSE -sOutputFile="
			#command = r"gswin64c.exe -dPrinted -q -sDEVICE=mswinpr2 -sPAPERSIZE="+paperType+" -dBATCH -dFitPage -dNOPROMPT -dFIXEDMEDIA -dNOPAUSE -sOutputFile="
			#the following string shoudl generate something like this: -sOutputFile="\\spool\KONICA MINOLTA 423"
			command = command + "\"\\\\spool\\"+printer+"\" "
			command = command + pdfPath
		#print (command)
	else:
		command = r"gswin64c.exe -dPrinted -q -sDEVICE=mswinpr2 -sPAPERSIZE="+paperType+" -dBATCH -dFitPage -dNOPROMPT -dFIXEDMEDIA -dNOPAUSE -sOutputFile="
		#command = r"gswin64c.exe -dPrinted -q -sDEVICE=mswinpr2 -dNoCancel -sPAPERSIZE="+paperType+" -dBATCH -dFitPage -dNOPROMPT -dFIXEDMEDIA -dNOPAUSE -sOutputFile="
		#the following string shoudl generate something like this: -sOutputFile="\\spool\KONICA MINOLTA 423"
		command = command + "\"\\\\spool\\"+printer+"\" "
		command = command + pdfPath
	"""
	command = r"gswin64c.exe -dPrinted -q -dNoCancel -sPAPERSIZE="+paperType+" -dBATCH -dFitPage -dNOPROMPT -dFIXEDMEDIA -dNOPAUSE "
	command = command + pdfPath
	"""
	filex.write(command+"\n")

	#print (command)


	try:
		gs = subprocess.Popen(command, shell=True, stdout=filex,stderr=subprocess.STDOUT,stdin=subprocess.PIPE)
	except:
		filex.write("POPEN\n")
		filex.write(str(sys.exc_info()[0]))
		filex.write(str(sys.exc_info()[1]))
		filex.write("\n")
	try:
		while True:
			gs.poll()
			if gs.returncode != None:
				#print ("Finished")
				return jobCounter + 1
	except:
		filex.write("returncode\n")
		filex.write(str(sys.exc_info()[0]))
		filex.write(str(sys.exc_info()[1]))
		filex.write("\n")


	filex.write("\n")
	f.close()

if __name__=="__main__":
	global isFrozen


	checkGhostScriptPath()
	root = tkinter.Tk()
	if getattr(sys, 'frozen', False):
		isFrozen = True
		root.iconbitmap(sys._MEIPASS+r"/emblem_print.ico")
	else:
		root.iconbitmap("emblem_print.ico")
		isFrozen = False
	root.style = ttk.Style()
	mainUI = mainUIClass(root)
	root.style.theme_use("winnative")

	try:
		f = open('C:\\Users\\'+os.getlogin()+'\\pdf_printer_settings','r+')
	except IOError:
		f = open('C:\\Users\\'+os.getlogin()+'\\pdf_printer_settings','w+')

	mainUI.selectedPrinter.set(f.readline().rstrip())
	f.close()
	root.mainloop()
