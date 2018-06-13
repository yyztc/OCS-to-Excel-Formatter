import csv
'''
Created on Jun 5, 2018

@author: DHorowitz
'''
class DataInput:
    def __init__(self, csvfile):
        self.csvfile = csvfile
        self.data = {}
        
    def getdata(self):
        return self.data

    def loaddata(self):
        with open(self.csvfile) as csvfile:
            reader = csv.DictReader(csvfile,delimiter=';')
            self.data = [r for r in reader]

# if __name__ == "__main__":
#     z = DataInput('export.csv')
#     z.loaddata()
#     data = z.getdata()