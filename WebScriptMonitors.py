import requests
from selenium import webdriver
from selenium.webdriver.support.ui import Select
'''
Created on Jun 6, 2018

@author: DHorowitz
'''

class WebScriptMonitors:
    def __init__(self, username, password, directory, systemid=0):
        self.username = username
        self.password = password
        self.directory = directory
        self.systemid = str(systemid)
        self.cookies = object
    
    def setCookies(self, cookies):
        self.cookies = cookies
    
    def getCookies(self):
        return self.cookies
    
    def grabCookies(self):
        options = webdriver.ChromeOptions()
        options.add_argument('headless')
        browser = webdriver.Chrome(chrome_options=options)
        browser.get('http://hqocssms01.mcnichols.com/ocsreports/')
        login = browser.find_element_by_name('LOGIN')
        passwd = browser.find_element_by_name('PASSWD')
        go = browser.find_element_by_name('Valid_CNX')
        
        login.send_keys(self.username)
        passwd.send_keys(self.password)
        go.click()
        browser.get('http://hqocssms01.mcnichols.com/ocsreports/index.php?function=visu_computers')
        select = Select(browser.find_element_by_name("restCollist_show_all"))
        select.select_by_value("Model")
        select2 = Select(browser.find_element_by_name("restCollist_show_all"))        
        select2.select_by_value("Computer Id")
        self.cookies = browser.get_cookies()        
        
        with requests.session() as r:
            payload = {
            'LOGIN': self.username,
            'PASSWD': self.password,
            'Valid_CNX' : 'Send'
            }
            r.post('http://hqocssms01.mcnichols.com/ocsreports/', data = payload)
            for cookie in self.cookies:
                r.cookies.set(cookie['name'], cookie['value'])
            r.get('http://hqocssms01.mcnichols.com/ocsreports/index.php?function=visu_computers')
            result = r.get("http://hqocssms01.mcnichols.com/ocsreports/index.php?function=export_csv&no_header=1&tablename=list_show_all&base=''")
            with open(self.directory, 'wb') as f:
                f.write(result.content)
     
    def OCS(self):
        with requests.session() as r:
            payload = {
            'LOGIN': self.username,
            'PASSWD': self.password,
            'Valid_CNX' : 'Send'
            }
            r.post('http://hqocssms01.mcnichols.com/ocsreports/index.php', data = payload)
            for cookie in self.cookies:
                r.cookies.set(cookie['name'], cookie['value'])
            r.get('http://hqocssms01.mcnichols.com/ocsreports/index.php?function=computer&head=1&systemid=' + self.systemid + '&option=cd_monitors')
            result = r.get("http://hqocssms01.mcnichols.com/ocsreports/index.php?function=export_csv&no_header=1&tablename=affich_monitors&base=''")
            with open(self.directory, 'wb') as f:
                f.write(result.content)
#          
# if __name__ == "__main__":
#     x = WebScriptMonitors('dhorowitz','science313', os.path.dirname(os.path.realpath(__file__)) + '/monitors.csv', 41087)
#     x.grabCookies()
#     z = ExcelGenerateMonitors('done.xlsx','export.csv')
#     z.grabData()
#     z.generate()