import xlsxwriter as xlsx
'''
Created on Jun 5, 2018

@author: DHorowitz
'''
class ExcelGenerateMonitors:
    def __init__(self, filename, data):
        self.filename = filename
        self.data = data
        self.longestString = []
        self.headerlist = ['User', 'Last Inventory', 'Computer', 'Monitor Manufacturer', 'Model/Notes', 'Serial Number']
    def getfilename(self):
        return self.filename
    
    def getdata(self):
        return self.data
    
    def getlongeststring(self):
        return self.longestString

    def generate(self):
        workbook = xlsx.Workbook(self.filename)
        worksheet = workbook.add_worksheet()

        keylooptemp = 0
        for key in self.headerlist:
            worksheet.write(0,keylooptemp, key)
            keylooptemp += 1
        temp = 1
        for elem in self.data:
            if len(self.data[elem]) > 1:
                for x in range(1, len(self.data[elem])):
                    colcounter = 0
                    worksheet.write(temp,colcounter, self.data[elem][0][0])
                    colcounter += 1
                    worksheet.write(temp,colcounter, self.data[elem][0][1])
                    colcounter += 1
                    worksheet.write(temp,colcounter, self.data[elem][0][2])
                    colcounter += 1
                    worksheet.write(temp, colcounter, self.data[elem][x][0])
                    colcounter += 1
                    worksheet.write(temp, colcounter, self.data[elem][x][1])
                    colcounter += 1
                    worksheet.write(temp, colcounter, self.data[elem][x][2])                                                                                                    
                    temp += 1
        worksheet.set_column(0, 5, 25)
        workbook.close()
        
# if __name__ == "__main__":
#     x = ExcelGenerateMonitors('done.xlsx', {'0' : [['jgardner', '2018', 'zzzzzzz'], ['Samsung', 'SyncMaster', 'HMDY702369']]})
#     x.generate()

    
