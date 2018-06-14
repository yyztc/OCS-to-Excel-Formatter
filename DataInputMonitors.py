'''
Created on Jun 12, 2018

@author: DHorowitz
'''
import csv
class DataInputMonitors:
    def __init__(self, csvfile):
        self.csvfile = csvfile
        self.monitordata = []
        
    def getmonitordata(self):
        return self.monitordata

    def loadmonitordata(self):
        with open(self.csvfile) as csvfile:
            reader = csv.reader(csvfile,delimiter=';')
            next(reader)
            for row in reader:
                if row[4] != '':
                    self.monitordata.append([row[0],row[1],row[4]])

# if __name__ == "__main__":
#     z = DataInputMonitors('monitors.csv')
#     z.loadmonitordata()
#     data = z.getmonitordata()
#     print(data)