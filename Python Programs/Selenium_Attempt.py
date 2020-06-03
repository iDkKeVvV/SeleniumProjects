from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import urllib.request
import xlrd 

workbook = xlrd.open_workbook('Asset_Inventory_Kohler.xlsx')
worksheet = workbook.sheet_by_name('Asset Inventory')
counter = 1301

while counter <= 1400:
    EMCOSKU = worksheet.cell(counter,0).value
    productnumber = worksheet.cell(counter,1).value
    counter= counter+1

    driver = webdriver.Chrome(executable_path=r"C:\Users\Kai\Desktop\chromedriver_win32\chromedriver.exe")
    driver.get('https://www.us.kohler.com/us/')
    elem = driver.find_element_by_xpath('//*[@id="nav-searchbox"]')
    elem.click()
    elem.send_keys(productnumber)
    elem.send_keys(Keys.RETURN)
    driver.get(driver.current_url)  

    current_url = 'https://www.us.kohler.com/us/s?Ntt=' + productnumber

    if driver.current_url == current_url:

        print (EMCOSKU)
        driver.close()
        pass 

    else:
        img = driver.find_element_by_xpath('//*[@id="heroImage"]')
        src = img.get_attribute('src')

        PNG = ".png"

        urllib.request.urlretrieve(src,(EMCOSKU + PNG))
        driver.close()
    
    

else:
    print("The test has been a success")
    pass