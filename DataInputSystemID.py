import csv
'''
Created on Jun 5, 2018

@author: DHorowitz
'''
class DataInputSystemID:
    def __init__(self, csvfile):
        self.csvfile = csvfile
        self.data = {}
        
    def getdata(self):
        return self.data

    def loaddata(self):
        with open(self.csvfile) as csvfile:
            reader = csv.reader(csvfile,delimiter=';')
            next(reader)
            for row in reader:
                self.data[row[8]] = [[row[3], row[1], row[2]]]

# if __name__ == "__main__":
#     z = DataInputSystemID('export.csv')
#     z.loaddata()
#     data = z.getdata()
#     print(data)