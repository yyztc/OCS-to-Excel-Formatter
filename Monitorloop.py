from ExcelGenerateMonitors import ExcelGenerateMonitors
from WebScriptMonitors import WebScriptMonitors
import os
from DataInputMonitors import DataInputMonitors
import time
from DataInputSystemID import DataInputSystemID
'''
Created on Jun 12, 2018

@author: DHorowitz
'''

class MonitorLoop:
    def __init__(self, username, password, finalfilename, directory):
        self.data = {}
        self.username = username
        self.password = password
        self.filename = finalfilename
        self.dir = directory
    def go(self):
        
        getcook = WebScriptMonitors(self.username, self.password, self.dir + '/monitors.csv')
        getcook.grabCookies()
        cookies = getcook.getCookies()
        
        z = DataInputSystemID('monitors.csv')
        z.loaddata()
        self.data = z.getdata()
        
        
        for key in self.data:
            if self.data[key] != ['']:
                try:
                    x = WebScriptMonitors(self.username, self.password, self.dir + '/monitors.csv', key)
                except:
                    time.sleep(5)
                    x = WebScriptMonitors(self.username, self.password, self.dir + '/monitors.csv', key)
                x.setCookies(cookies)
                x.OCS()
                z = DataInputMonitors('monitors.csv')
                z.loadmonitordata()
                tempdata = z.getmonitordata()
                if tempdata != []:
                    self.data[key]= self.data[key] + tempdata
    
        c = ExcelGenerateMonitors(self.filename, self.data)
        c.generate()      
if __name__ == "__main__":
    z = MonitorLoop('dhorowitz', 'science313', 'done.xlsx')
    z.go()