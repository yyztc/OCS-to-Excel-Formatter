from pandas import read_excel as pr
from pandas import ExcelWriter as pv
from ExcelGenerate import ExcelGenerate
from WebScript import WebScript

'''
Created on Jun 7, 2018

@author: DHorowitz
'''
class ExcelFilter:
    def __init__(self, filename):
        self.ramList = []
        self.userList = []
        self.osList = []
        self.year = []
        self.branch = []
        self.model = []
        self.filename = filename
        self.df = pr(self.filename, sheet_name='Sheet1')
   
    def getramlist(self):
        return self.ramList
    
    def getuserlist(self):
        return self.userList
    
    def getoslist(self):
        return self.osList
   
    def getyearlist(self):
        return self.year
    
    def getbranchlist(self):
        return self.branch
    
    def getmodellist(self):
        return self.model
   
    def getLists(self):
        if 'RAM (MB)' in self.df:
            self.ramList = list(self.df['RAM (MB)'].unique())
            self.ramList.sort()
        if 'Computer' in self.df:
            temp = list(self.df['Computer'].unique())
            for item in temp:
                if '-' in item:
                    if item[0:item.find('-')] not in self.branch and (len(item[0:item.find('-')]) == 3 or len(item[0:item.find('-')]) == 2):
                        self.branch.append(item[0:item.find('-')])
            self.branch.sort()
        if 'User' in self.df:
            self.userList = list(self.df['User'].unique())
        if 'Operating system' in self.df:
            self.osList = list(self.df['Operating system'].unique())
            self.osList.sort()
        if 'Last inventory' in self.df:
            temp = list(self.df['Last inventory'])
            for item in temp:
                if item[0:4] not in self.year:
                    self.year.append(item[0:4])
        if 'Model' in self.df:
            self.model = list(self.df['Model'].unique())
            for x in range(0,len(self.model)):
                self.model[x] = str(self.model[x])
            self.model.sort()


    def collectColumns(self, data, longestString):
        if 'Account info: TAG' in self.df:
            self.df = self.df.drop('Account info: TAG', 1)

        writer = pv(self.filename, engine='xlsxwriter')
        self.df.to_excel(writer, index=False)
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        for col in range(len(data)):
            worksheet.set_column(col, col, len(longestString[col]) *1.1)
        
        workbook.close()
         
    def dropram(self, ramval):
        self.df = self.df.drop(self.df[self.df['RAM (MB)'] != ramval].index)
        
    def dropuser(self, userval):
        self.df = self.df.drop(self.df[self.df['User'] != userval].index)
        
    def dropyear(self, year):
        self.df = self.df.drop(self.df[self.df['Last inventory'].str[0:4] != year].index)
        
    def dropos(self, osval):
        self.df = self.df.drop(self.df[self.df['Operating system'] != osval].index)
        
    def dropbranch(self, branch):
        if len(branch) == 2:
            tempbranch = branch + '-'
            self.df = self.df.drop(self.df[self.df['Computer'].str[0:3] != tempbranch].index)
        elif len(branch) == 3:
            tempbranch = branch + '-'
            self.df = self.df.drop(self.df[self.df['Computer'].str[0:4] != tempbranch].index)
        else:
            pass       
        
    def dropmodel(self, model):
        self.df = self.df.drop(self.df[self.df['Model'] != model].index)

# if __name__ == "__main__":
#     x = WebScript('dhorowitz','science313')
#     x.OCS()
#     z = ExcelGenerate('done.xlsx','export.csv')
#     z.grabData()
#     z.generate()
#     
#     
#     q = ExcelFilter()
#     q.getLists()
#     q.dropbranch('ATL')
#     q.collectColumns(z.getdata(),z.getlongeststring())