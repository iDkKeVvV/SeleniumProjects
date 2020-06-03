from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from xlutils import copy
from xlwt import Workbook
import urllib.request
import xlrd 
import xlwt
import time 

workbook = xlrd.open_workbook('Asset_Inventory_Kohler.xlsx')
worksheet = workbook.sheet_by_name('Asset Inventory')

NOSKU = Workbook()
sheet = NOSKU.add_sheet('Sheet 1')

counter = 2300
start_time = time.time()

while counter <= 2500:
    EMCOSKU = worksheet.cell(counter,0).value
    productnumber = worksheet.cell(counter,1).value
    
    driver = webdriver.Chrome(executable_path=r"C:\Users\Kai\Desktop\chromedriver_win32\chromedriver.exe")
    driver.get('https://www.us.kohler.com/us/')
    try:
        counter= counter+1
        elem = driver.find_element_by_xpath('//*[@id="nav-searchbox"]')
        elem.click()
        elem.send_keys(productnumber)
        elem.send_keys(Keys.RETURN)
        driver.get(driver.current_url)  

        current_url = 'https://www.us.kohler.com/us/s?Ntt=' + productnumber

        if driver.current_url == current_url:
            sheet.write(counter,16,'NO')
            sheet.write(counter,0,EMCOSKU)
            print (EMCOSKU)
            driver.close()
            pass 

        else:
            img = driver.find_element_by_xpath('//*[@id="heroImage"]')
            src = img.get_attribute('src')

            PNG = ".png"

            urllib.request.urlretrieve(src,(EMCOSKU + PNG))
            driver.close()

    except NoSuchElementException:
        driver.close()
        pass
            
else:
    NOSKU.save('No_ImagesTrial.xls')
    elapsed_time = (time.time() - start_time)
    print(str(elapsed_time) + " seconds")
    pass
