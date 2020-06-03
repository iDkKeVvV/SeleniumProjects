import urllib.request
from selenium import webdriver
driver = webdriver.Chrome(executable_path=r"C:\Users\Kai\Desktop\chromedriver_win32\chromedriver.exe")
driver.get('https://www.us.kohler.com/us//productDetail/serviceparts:627711/627711.htm?skuId=592264&brandId=empty&')

img = driver.find_element_by_xpath('//*[@id="heroImage"]')
src = img.get_attribute('src')

urllib.request.urlretrieve(src,"test.png")

driver.close()
