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
jsonpath = 'cap_rongyu/' + str(today)

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
ori_data = {'time': now}
for i in range(1, 100):
    time.sleep(1)
    br.get('https://console.captainbi.com/FBA-fba_inventory_v3-fba_stock_asin_list.html?m=FBA&c=fba_inventory_v3&a'
           '=fba_asin_list&isajax=1&is_list=1&page= ' + str(i))
    time.sleep(1)
    json_text = br.find_element(By.CSS_SELECTOR, 'pre').get_attribute('innerText')
    json_response = json.loads(json_text)
    for row in json_response['rows']:
        if row['country_code'] != 'US':
            continue
        ori_data[row['seller_sku']] = row
    page_total = json_response['records']/20 + 2
    if i > page_total:
        break



# br.get('https://console.captainbi.com/amzcaptain-amazon_finance_operate-finance_add_keywords_product_analysis.html')
# time.sleep(1)
# br.find_element(By.XPATH, '//*[@id="site-select"]').click()
# time.sleep(1)
# br.find_element(By.XPATH, '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div/div/div/ul[2]/li[2]/b').click()
# time.sleep(1)
# keys = br.find_element(By.XPATH, '//*[@id="list"]/tbody/tr[1]/td[4]').text.split('\n')
# poss = br.find_element(By.XPATH, '//*[@id="list"]/tbody/tr[1]/td[6]').text.split('\n')
#
# keywords = {}
# for idx, key in enumerate(keys):
#     if key == '+添加关键词':
#         break
#     pos = poss[idx].replace('页，第', '-').replace('第', 'P').replace('个', '')
#     keywords[key] = pos


# ori_data['keywords'] = keywords
ori_data['keywords'] = {}


with open(jsonpath, 'w') as fp:
    json.dump(ori_data, fp)

# br.close()