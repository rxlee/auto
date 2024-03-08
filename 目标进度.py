from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import json
import time
import random

time.sleep(random.randint(5, 10))
''' 爬取船长库存销量数据 '''

now = int(time.time())
today = time.strftime("%Y%m%d", time.localtime())
jsonpath = 'cap/' + str(today)

# chrome_opt = webdriver.ChromeOptions()
# chrome_opt.add_experimental_option("excludeSwitches", ["ignore-certificate-errors"])
# chrome_opt.add_argument('--disable-gpu')
# chrome_opt.add_argument('--allowed-origins')
# path = r"C:/Users/Administrator/chromedriver.exe"
# br = webdriver.Chrome(executable_path=path, chrome_options=chrome_opt)
#

options = Options()
options.add_experimental_option('debuggerAddress', '127.0.0.1:56743')
service = webdriver.chrome.service.Service(r"C:/Users/Administrator/chromedriver.exe")
br = webdriver.Chrome(service=service, options=options)
br.get('https://console.captainbi.com/app/#/main/homePage')
time.sleep(3)
try:
    conti = br.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/button[2]')
    br.execute_script('arguments[0].click()', conti)
    time.sleep(1)
    br.find_element(By.ID, 'username').send_keys(Keys.CONTROL, 'a')
    br.find_element(By.ID, 'username').send_keys(Keys.BACKSPACE)
    time.sleep(1)
    br.find_element(By.ID, 'username').send_keys('Liruixiang666')
    time.sleep(1)
    br.find_element(By.ID, 'password').send_keys(Keys.CONTROL, 'a')
    br.find_element(By.ID, 'password').send_keys(Keys.BACKSPACE)
    time.sleep(1)
    br.find_element(By.ID, 'password').send_keys('aA10547')
    time.sleep(1)
    submit = br.find_element(By.ID, 'submit')
    br.execute_script('arguments[0].click()', submit)
except Exception as e:
    pass

time.sleep(3)