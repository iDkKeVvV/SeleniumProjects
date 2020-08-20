from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from pathlib import Path
import os 
import shutil
import xlrd 
import xlwt
import time 
from xlwt import Workbook
import os
import shutil

workbook = xlrd.open_workbook('Asset_Inventory_Kohler.xlsx')
worksheet = workbook.sheet_by_name('Asset Inventory')
xlcolumn = 0

NOSKU = Workbook()
sheet = NOSKU.add_sheet('Sheet 1')

counter = 10000
start_time = time.time()
delay = 3

while counter <= 10200:
    EMCOSKU = worksheet.cell(counter,0).value
    productnumber = worksheet.cell(counter,1).value

    options = Options()
    options.add_experimental_option("prefs", {"download.prompt_for_download": False, "plugins.always_open_pdf_externally": True})
    driver = webdriver.Chrome(executable_path=r"C:\Users\Kai\Documents\chromedriver_win32\chromedriver.exe", options = options)
    wait = WebDriverWait(driver, 10)
    driver.get('https://www.us.kohler.com/us/')

    counter= counter + 1
    myElem = WebDriverWait(driver,delay).until(EC.presence_of_element_located((By.XPATH,'//*[@id="nav-searchbox"]')))
    elem = driver.find_element_by_xpath('//*[@id="nav-searchbox"]')
    elem.click()
    elem.send_keys(productnumber)
    elem.send_keys(Keys.RETURN)
    driver.get(driver.current_url)    

    current_url = 'https://www.us.kohler.com/us/s?Ntt=' + productnumber

    try:
        elements = driver.find_elements_by_partial_link_text("Rough In/Spec Sheet")

        if not elements:
            print("TEXT NOT FOUND, Moving Into second loop")
            elements2 = driver.find_elements_by_partial_link_text("Specification Sheet")

            if not elements2:
                print("You have been jebaited")
                sheet.write(xlcolumn,16,'NO')
                sheet.write(xlcolumn,0,EMCOSKU)
                xlcolumn = xlcolumn + 1
                print (EMCOSKU)
                driver.quit()

            else:
                element2 = elements2[0]
                driver.execute_script("arguments[0].click();", element2)
                spec_extension = "_specs.pdf"
                time.sleep(2)
                New_Name = EMCOSKU + spec_extension
                filepath = r'C:\Users\Kai\Downloads'
                filename = max([filepath +"\\"+ f for f in os.listdir(filepath)], key=os.path.getctime)
                shutil.move(os.path.join(r'C:\Users\Kai\Downloads',filename),New_Name)
                #filename = max([r"C:\Users\Kai\Downloads" + "\\" + f for f in os.listdir(r"C:\Users\Kai\Downloads")],key=os.path.getctime)
                #shutil.move(filename,os.path.join(r"C:\Users\Kai\Downloads",(EMCOSKU + spec_extension)))
                #print("The file has been saved")
                driver.quit()
            
        else:
            element = elements[0]
            driver.execute_script("arguments[0].click();", element)
            spec_extension = "_specs.pdf"
            time.sleep(2)
            New_Name = EMCOSKU + spec_extension
            filepath = r'C:\Users\Kai\Downloads'
            filename = max([filepath +"\\"+ f for f in os.listdir(filepath)], key=os.path.getctime)
            shutil.move(os.path.join(r'C:\Users\Kai\Downloads',filename),New_Name)
            #filename = max([r"C:\Users\Kai\Downloads" + "\\" + f for f in os.listdir(r"C:\Users\Kai\Downloads")],key=os.path.getctime)
            #shutil.move(filename,os.path.join(r"C:\Users\Kai\Downloads",(EMCOSKU + spec_extension)))
            #print("The file has been saved")
            driver.quit()
            

    except NoSuchElementException:
        print("NoSuchElementException occured")
        sheet.write(xlcolumn,16,'NO')
        sheet.write(xlcolumn,0,EMCOSKU)
        xlcolumn = xlcolumn + 1
        print (EMCOSKU)
        print("The data has been added to the excel sheet")
        driver.quit()
            

else:
    print("The data set has been completed successfully")
    NOSKU.save('No_PDF.xls')
    elapsed_time = (time.time() - start_time)
    print(str(elapsed_time) + " seconds")
    
