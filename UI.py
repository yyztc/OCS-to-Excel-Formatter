'''
Created on Jun 7, 2018

@author: DHorowitz
'''
import tkinter
from tkinter import *
from tkinter import messagebox
from ExcelGenerate import ExcelGenerate
from WebScript import WebScript
from ExcelFilter import ExcelFilter
import os.path
from datetime import date
from Monitorloop import MonitorLoop

class MyFirstGUI(Frame):
    def __init__(self, master):
        self.excelFilt = object
        self.excelGen = object
        
        self.curryear = str(date.today().year)
        self.lstyear = str(date.today().year - 1)
        
        self.cookies=object
        
        self.master = master
        super().__init__(master)
        master.title("OCS Inventory Excel Grabber")

        self.dir = os.path.dirname(os.path.realpath(__file__))
        self.filename = "export.csv"
        self.exportfilename = ''

        self.ram = ''
        self.os = ''
        self.branch = ''
        self.year = ''
        self.user = ''
        self.model = ''
        
        self.label = Label(master, text="OCS Inventory Excel Grabber")
        self.label.pack()
        
        self.label_username = Label(self, text="Username")
        self.label_password = Label(self, text="Password")
        self.label_filename = Label(self, text='File name')

        self.entry_username = Entry(self)
        self.entry_password = Entry(self, show="*")
        self.entry_filename = Entry(self)
        
        self.entry_filename.delete(0, END)
        self.entry_filename.insert(0, "Done")
        
        self.label_username.grid(row=0, sticky=E)
        self.label_password.grid(row=1, sticky=E)
        self.label_filename.grid(row=2, sticky=E)
        self.entry_username.grid(row=0, column=1)
        self.entry_password.grid(row=1, column=1)
        self.entry_filename.grid(row=2, column=1)
        
        self.logbtn = Button(self, text="Login and Go", command=self.run)
        self.logbtn.grid(columnspan=2, pady = 10)
        master.bind('<Return>', self.func)
        
        
        self.monitorlabel = Label(self, text="Clicking below takes a few minutes.")
        self.monitorbrn = Button(self, text="Login and Launch Monitor Grabber", command=self.run2)
        self.monitorbrn.grid(row = 5, columnspan = 2, pady=10, padx = 10)
        self.monitorlabel.grid(row=4, columnspan=2)
        
        self.pack()
        

    def getDir(self):
        return self.dir

    def func(self, something):
        self.run()
    
    def quit(self):
        global root
        root.destroy()

    def run2(self):
        self.exportfilename = self.entry_filename.get() +'.xlsx'
        if os.path.isfile(self.exportfilename):
            result = messagebox.askokcancel(title='Overwrite?', message='The file "' +  self.exportfilename + '" is about to be overwritten. Proceed?')
            if result == True:
                pass
            else:
                return self.master
        self.excelGen = ExcelGenerate(self.exportfilename,self.filename)
        try:
            x = WebScript(str(self.entry_username.get()),str(self.entry_password.get()), self.dir + '/' + self.filename)
            x.OCS()
            self.excelGen.grabData()  
            self.excelGen.generate()
            os.remove(self.exportfilename)
        except:
            messagebox.showerror("Error", "Invalid Login")
            return self.master
        
        os.remove('export.csv')
        run = MonitorLoop(str(self.entry_username.get()),str(self.entry_password.get()), self.exportfilename, self.dir)
        run.go()
        os.remove('monitors.csv')
        messagebox.showinfo('Complete', 'Monitor Spreadsheet has been constructed.')
        self.quit()

    def run(self):
        self.exportfilename = self.entry_filename.get() +'.xlsx'
        if os.path.isfile(self.exportfilename):
            result = messagebox.askokcancel(title='Overwrite?', message='The file "' +  self.exportfilename + '" is about to be overwritten. Proceed?')
            if result == True:
                pass
            else:
                return self.master
        self.excelGen = ExcelGenerate(self.exportfilename,self.filename)
        try:
            x = WebScript(str(self.entry_username.get()),str(self.entry_password.get()), self.dir + '/' + self.filename)
            x.OCS()
            self.excelGen.grabData()  
            self.excelGen.generate()
        except:
            messagebox.showerror("Error", "Invalid Login")
            return self.master
        
        self.filterwin()
        os.remove('export.csv')


    def filterwin(self): # new window definition
        global root 
        root.withdraw()
        
        
        self.excelFilt = ExcelFilter(self.exportfilename)
        self.excelFilt.getLists(self.curryear, self.lstyear)
       
        
        newwin = Toplevel(root)
        frame = Frame(newwin)
        newwin.title("Filter the results")
        label = Label(newwin, text="Filter the results. Click finish after completion.")
        label.pack()
        
        oslabel = Label(frame, text = "Operating System")
        osvar=StringVar(frame)
        osoptions = OptionMenu(frame, osvar, *self.excelFilt.getoslist(), command=self.osboolset)
        
        branchlabel = Label(frame, text = "Branch")
        branchvar=StringVar(frame)
        branchoptions = OptionMenu(frame, branchvar, *self.excelFilt.getbranchlist(), command=self.branchboolset)
        
        yearlabel = Label(frame, text = "Year")
        yearvar=StringVar(frame)
        yearoptions = OptionMenu(frame, yearvar, *self.excelFilt.getyearlist(), command=self.yearboolset)

        modellabel = Label(frame, text = "Model")
        modelvar=StringVar(frame)
        modeloptions = OptionMenu(frame, modelvar, *self.excelFilt.getmodellist(), command=self.modelboolset)

        oslabel.grid(row=0, sticky=E)
        branchlabel.grid(row=1, sticky=E)
        yearlabel.grid(row=2, sticky=E)
        modellabel.grid(row=3, sticky=E)        
        osoptions.grid(row=0, column=1)
        branchoptions.grid(row=1, column=1)
        yearoptions.grid(row=2, column=1)
        modeloptions.grid(row=3, column=1)
        
        frame.pack()
        
        logbtn = Button(newwin, text="Finish", command=self.runfilter)
        logbtn.pack()
        newwin.bind('<Return>', self.func2)
        

    
    def osboolset(self, val):
        self.os = val
    
    def ramboolset(self, val):
        self.ram = val
    
    def branchboolset(self, val):
        self.branch = val
    
    def yearboolset(self, val):
        self.year = val
        
    def userboolset(self, val):
        self.user = val
    
    def modelboolset(self, val):
        self.model = val

    def runfilter(self):
        if self.ram != '':
            self.excelFilt.dropram(self.ram)
        if self.os != '':
            if self.os == 'Microsoft Windows 7 (ALL Versions)':
                self.excelFilt.dropos('Microsoft Windows 7 Enterprise', 'Microsoft Windows 7 Professional')
            elif self.os == 'Microsoft Windows 10 (ALL Versions)':
                self.excelFilt.dropos('Microsoft Windows 10 Enterprise', 'Microsoft Windows 10 Pro', 'Microsoft Windows 10 Enterprise 2016 LTSB')
            elif self.os == 'Microsoft Windows 8.1 (ALL Versions)':
                self.excelFilt.dropos('Microsoft Windows 8.1 Enterprise', 'Microsoft Windows 8.1 Pro')
            else:
                self.excelFilt.dropos(self.os)
        if self.branch != '':
            self.excelFilt.dropbranch(self.branch)
        if self.year != '':
            if self.year == (self.curryear + ' & ' + self.lstyear):
                self.excelFilt.dropyear(self.curryear, self.lstyear)
            else:
                self.excelFilt.dropyear(self.year)
        if self.user != '':
            self.excelFilt.dropuser(self.user)
        if self.model != '':
            self.excelFilt.dropmodel(self.model)
        
        self.excelFilt.collectColumns(self.excelGen.getdata(),self.excelGen.getlongeststring())
        self.quit()
    
    def func2(self,something):
        self.runfilter()
    
root = Tk()
my_gui = MyFirstGUI(root)
root.mainloop()
