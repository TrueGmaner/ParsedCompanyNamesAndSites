# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
# Press Ctrl+F8 to toggle the breakpoint.
import string

import xlrd
from selenium import webdriver
from xlutils.copy import copy  # http://pypi.python.org/pypi/xlutils
from selenium.webdriver.chrome.service import Service


def readExcelTable():
    r_Book = xlrd.open_workbook('CompanyNames.xls', formatting_info=True)
    r_Sheet = r_Book.sheet_by_index(0)
    w_Book = copy(xlrd.open_workbook('CompanyNames.xls'))
    w_Sheet = w_Book.get_sheet(0)

    for i in range(len(w_Sheet.rows)):
        print(r_Sheet.cell_value(i, 0))
        companyName = r_Sheet.cell_value(i, 0)

        options = webdriver.ChromeOptions()
        options.add_argument('headless')
        service = Service('C:/Program Files/ChromeDriver/chromedriver.exe')
        driver = webdriver.Chrome(service=service, options=options)

        driver.get('https://www.google.com/search?q=' + companyName + '+official+site')
        driver.implicitly_wait(20)
        xpathsList = ['//*[@id="rso"]/div[1]/div/div/div[1]/div/div/div[1]/div/a/div/cite', '//*[@id="rso"]/div[1]/div/div/div[1]/div/a/div/cite']

        for xpath in xpathsList:
            try:
                siteName = driver.find_element("xpath", xpath)
            except:
                print("exception caught")
            else:
                break

        siteName = str(siteName.text.split()[0])
        index = siteName.find('.') + 1
        siteName = siteName[index:]
        w_Sheet.write(i, 1, siteName)

    driver.quit()
    w_Book.save('CompanyNames.xls')


readExcelTable()










# class Data:
#     def __init__(self, name):
#         self.name = name
#         self.url = None
#
# def readExcelTable():
#     r_Book = xlrd.open_workbook('CompanyNames.xls', formatting_info=True)
#     r_Sheet = r_Book.sheet_by_index(0)
#     dataArr = []
#
#     for i in range(0, 6):
#         dataArr.append(Data(r_Sheet.cell_value(i, 0)))
#         print(dataArr[i].name)
#     return dataArr
#
# def convertData(dataArr):
#     for i in range(0, (len(dataArr))):
#         (Data)(dataArr[i]).url = "123"
#         print(dataArr[i].url)
#     return dataArr
#
# def writeData(dataArr):
#     w_Book = xlwt.Workbook()
#     sheet1 = w_Book.add_sheet('Лист1') #type: Workbook
#
#     for i in range(0, 6):
#         w_Book.get_sheet(0).write(i, 0, dataArr[i].name)
#         w_Book.get_sheet(0).write(i, 1, dataArr[i].url)
#     w_Book.save('CompanyNames.xls')
#
# dataArr = readExcelTable()
# dataArr2 = convertData(dataArr)
# writeData(dataArr2)