from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import urllib.request
import xlrd 

workbook = xlrd.open_workbook('Asset_Inventory_Kohler.xlsx')
worksheet = workbook.sheet_by_name('Asset Inventory')
counter= 1

while counter<= 4:
    EMCOSKU = (worksheet.cell(counter,0).value)
    value = (worksheet.cell(counter,1).value)
    print(value)
    print(EMCOSKU)
    counter= counter+1

else:
    print("The test has been a success")
    pass