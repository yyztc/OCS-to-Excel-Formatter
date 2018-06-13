import xlsxwriter as xlsx
'''
Created on Jun 5, 2018

@author: DHorowitz
'''
from DataInput import DataInput
class ExcelGenerate:
    def __init__(self, filename, csvfile):
        self.filename = filename
        self.csvfile = csvfile
        self.data = {}
        self.longestString = []
    
    def getfilename(self):
        return self.filename
    
    def getdata(self):
        return self.data
    
    def getlongeststring(self):
        return self.longestString
    
    def grabData(self):
        loader = DataInput(self.csvfile)
        loader.loaddata()
        self.data = loader.getdata()

    def generate(self):
        workbook = xlsx.Workbook(self.filename)
        worksheet = workbook.add_worksheet()
        self.longestString = [''] * len(self.data)
        keylooptemp = 0
        for key in self.data[0]:
            worksheet.write(0,keylooptemp, key)
            if len(self.longestString[keylooptemp-1]) < len(key):
                self.longestString[keylooptemp-1] = key
            keylooptemp+=1
        temp = 0
        for elem in self.data:
            colcounter = 0
            for key in elem:
                if (self.data[temp][key].isdigit()):
                    worksheet.write_number(temp+1,colcounter, int(self.data[temp][key]))
                else:
                    worksheet.write(temp+1,colcounter, self.data[temp][key])
                if len(self.longestString[colcounter-1]) < len(self.data[temp][key]):
                    self.longestString[colcounter-1] = self.data[temp][key]
                colcounter+=1
            temp += 1
        
        workbook.close()
        
# if __name__ == "__main__":
#     x = ExcelGenerate('done.xlsx','export.csv')
#     x.getData()
#     x.generate()

    
