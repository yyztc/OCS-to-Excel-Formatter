import requests
import os
from ExcelGenerate import ExcelGenerate 
from selenium import webdriver
from selenium.webdriver.support.ui import Select
'''
Created on Jun 6, 2018

@author: DHorowitz
'''

class WebScript:
    def __init__(self, username, password, directory):
        self.username = username
        self.password = password
        self.directory = directory

    def OCS(self):
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
        cookies = browser.get_cookies()
        
        with requests.session() as r:
            payload = {
            'LOGIN': self.username,
            'PASSWD': self.password,
            'Valid_CNX' : 'Send'
            }
            r.post('http://hqocssms01.mcnichols.com/ocsreports/index.php?function=visu_computers', data = payload)
            for cookie in cookies:
                r.cookies.set(cookie['name'], cookie['value'])
            r.post('http://hqocssms01.mcnichols.com/ocsreports/index.php?function=visu_computers')
            result = r.get("http://hqocssms01.mcnichols.com/ocsreports/index.php?function=export_csv&no_header=1&tablename=list_show_all&base=''")
            with open(self.directory, 'wb') as f:
                f.write(result.content)


# if __name__ == "__main__":
#     x = WebScript('dhorowitz','science313', os.path.dirname(os.path.realpath(__file__)) + '/export.csv')
#     x.OCS()
#     z = ExcelGenerate('done.xlsx','export.csv')
#     z.grabData()
#     z.generate()