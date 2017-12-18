# -*- coding: utf-8 -*-
from Tkinter import *
import tkFileDialog
import csv
import xlrd
import xlwt
import os
from time import gmtime,strftime
class App():
	#建立GUI界面
	def __init__(self, root):
		self.NL = 0
		self.RD = 1
		self.frm = Frame(root,width = 800, height = 600)
		self.frm.pack(fill = 'both',expand = True)
		
		self.menubar = Menu(root)
		self.menubar.add_command(label="calculate", command = self.calculation)
		self.menubar.add_command(label='help',command = self.help)
		root.config(menu = self.menubar)
		
		
		self.nameButton = Button(self.frm, text = 'namelist', command = self.getnamelist)
		self.nameButton.grid(row = 0, column = 0)
		self.dataButton = Button(self.frm, text = 'rawdata', command = self.getrawdata)
		self.dataButton.grid(row = 0,column = 2)
		
		self.nametext = Text(self.frm,width = 15)
		self.nametext.grid(row = 1,column = 0, sticky = 'nsew')
		self.datatext = Text(self.frm)
		self.datatext.grid(row = 1,column = 2, sticky = 'nsew')
		
		self.namescrb = Scrollbar(self.frm, command = self.nametext.yview)
		self.namescrb.grid(row = 1, column = 1, sticky = 'nsew')
		self.nametext['yscrollcommand'] = self.namescrb.set
		self.datascrb = Scrollbar(self.frm, command = self.datatext.yview)
		self.datascrb.grid(row = 1, column = 3, sticky = 'nsew')
		self.datatext['yscrollcommand'] = self.datascrb.set
	
	#打开文件对话框	
	def openfile(self):
		return tkFileDialog.askopenfilename()
		
	def createfiledir(self):
		self.newfolder = os.path.dirname(os.path.realpath(__file__)) + "\\" + strftime("%Y%m%d",gmtime())
		try:
			os.makedirs(self.newfolder)
		except Exception:
			print "cannot create new folder: ",self.newfolder


	#由Button触发,选取excel文件,转换为csv文件,并显示在文本框里
	def exceltocsv(self, csvname, id):
		self.theBook = xlrd.open_workbook(self.openfile())
		self.theSheet = self.theBook.sheet_by_index(0)
		
		self.createfiledir()
		with open(self.newfolder +'\\'+csvname,'wb') as csvfile:
			self.theWriter = csv.writer(csvfile,quoting = csv.QUOTE_ALL)
			for row in xrange(self.theSheet.nrows):
				self.theWriter.writerow([unicode(entry).encode("utf-8") for entry in self.theSheet.row_values(row)])
		#将CSV文件输出到文本框中
		with open(self.newfolder +'\\'+csvname,'rb') as csvfile:
			self.theReader = csv.reader(csvfile)
						
			if id == 0:
				self.nametext.delete('1.0',END)
				for rows in self.theReader:
					self.nametext.insert(INSERT,rows[0]+'\n')
			elif id == 1:
				self.datatext.delete('1.0',END)
				for rows in self.theReader:
					for cell in rows:
						self.datatext.insert(INSERT, cell + ',')
					self.datatext.insert(INSERT,'\n')
			else: 
				print 'cannot identify csvfile type'
	
	#将文本框修改过的文本储存于原csv文件
	def csvmodify(self, datacsv, namelscsv):
		with open(self.newfolder+'\\'+datacsv,'wb') as self.newcsv:
			self.modiwriter = csv.writer(self.newcsv,quoting = csv.QUOTE_ALL)
			self.alist = self.datatext.get(1.0,END).split('\n')
			for row in self.alist:
				self.modiwriter.writerow(row.split(','))
		with open(self.newfolder+'\\'+namelscsv,'wb') as self.newcsv:
			self.modiwriter = csv.writer(self.newcsv,quoting = csv.QUOTE_ALL)
			self.alist = self.nametext.get(1.0,END).split('\n')
			for row in self.alist:
				self.modiwriter.writerow(row.split(','))
	
	#运行csvmodify保存对csv文件的修改，计算生成最终统计文件
	def calculation(self):
		self.namelist = []
		self.medict = {}
		self.csvmodify(datacsv = self.datacsv, namelscsv = self.namelscsv)
		
		with open(self.newfolder +'\\'+ self.namelscsv,'rb') as namefile:
			self.namereader = csv.reader(namefile)
			for row in self.namereader:
				self.namelist.append(row[0])
			#print self.namelist
		
		with open(self.newfolder +'\\'+ self.datacsv,'rb') as sourcefile:
			self.sourcereader = csv.reader(sourcefile)
			for row in self.sourcereader:
				try:
					if row[2] in self.namelist:
						if row[3] in self.medict:
							self.medict[row[3]] = self.medict[row[3]] + float(row[4])
						else:
						# print row[4].decode('utf-8')
							self.medict[row[3]] = float(row[4])
				except Exception:
					print 'the calculation encounted some problems!'
		with open(self.newfolder +'\\'+'finaldata'+ strftime("%Y%m%d",gmtime()) +'.csv','wb') as self.medtotal:
			for i in self.medict:
				self.medtotal.write(i + ',' + str(self.medict[i]) + '\n')
				
	def help(self):
		print 'help content'
	
	#运行exceltocsv处理namelist
	def getnamelist(self):
		self.namelscsv = 'namelist'+ strftime("%Y%m%d",gmtime()) +'.csv'
		self.exceltocsv(self.namelscsv, self.NL)
	
	#运行exceltocsv处理原始数据rawdata
	def getrawdata(self):
		self.datacsv = 'rawdata'+ strftime("%Y%m%d",gmtime()) +'.csv'
		self.exceltocsv(self.datacsv, self.RD)
		
if __name__ == '__main__':
	reload(sys)
	sys.setdefaultencoding('utf-8')
	tk = Tk()
	app = App(tk)
	tk.mainloop()
	
