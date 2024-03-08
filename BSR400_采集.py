from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import json
import time
import os
import random
time.sleep(random.randint(5, 10))

dir_json = 'bsr400/'

now = int(time.time())
t = time.strftime("%Y-%m-%d %H_%M_%S", time.localtime())
json_file = dir_json + str(t) + '.json'



options = Options()
options.add_experimental_option('debuggerAddress', '127.0.0.1:56743')
service = webdriver.chrome.service.Service(r"C:/Users/Administrator/chromedriver.exe")
br = webdriver.Chrome(service=service, options=options)

all = {'time': now}

data = []
def gather(url):
    br.get(url)
    br.maximize_window()
    time.sleep(random.randint(2, 5))
    br.execute_script('window.scrollTo(0,3500)')
    time.sleep(1)
    br.execute_script('window.scrollTo(0,4500)')
    time.sleep(1)
    br.execute_script('window.scrollTo(0,5500)')
    time.sleep(1)
    br.execute_script('window.scrollTo(0,6500)')
    time.sleep(1)
    j = json.loads(br.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]').get_attribute('data-client-recs-list'))
    nodes = br.find_elements(By.ID, 'gridItemRoot')
    for index, node in enumerate(nodes):
        idx = index + 1
        print(str(idx)+' start ...')
        bsr = 401
        try:
             bsr = br.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div['+str(idx)+']/div/div[1]/div[1]/span').text.replace('#', '').replace(',', '')
        except:
            print('error-下架')
            continue
        title = br.find_element(By.XPATH,
                              '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[' + str(
                                  idx) + ']/div/div[2]/div/a[2]/span/div').text
        score = 0
        rating = 0
        price = 0
        try:
            price = float(br.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[' + str(idx) + ']/div/div[2]/div/div[2]/a').text.split('$')[1].replace(',', ''))
        except:
            try:
                price = float(br.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[' + str(idx) + ']/div/div[2]/div/div/a').text.split('$')[1].replace(',', ''))
            except:
                try:
                    price = float(br.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[' + str(idx) + ']/div/div[2]/div/div[2]/div/div/a').text.split('$')[1].replace(',', ''))
                except:
                    print('error-获取价格出错')
                    pass
                pass
            pass
        try:
            score = br.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[' + str(idx) + ']/div/div[2]/div/div[1]/div/a').get_attribute('title').split()[0]
        except:
            print('error-获取评分出错')
            pass
        try:
            rating = int(br.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div/div[2]/div[1]/div[1]/div[' + str(idx) + ']/div/div[2]/div/div[1]/div/a/span').text.replace(',', ''))
        except:
            print('error-获取rating出错')
            pass
        for item in j:
            if item['metadataMap']['render.zg.rank'] == bsr:
                item['bsr'] = bsr
                item['title'] = title
                item['score'] = score
                item['rating'] = rating
                item['price'] = price
        print(str(idx)+' end ...')
    data.extend(j)

gather('https://www.amazon.com/gp/bestsellers/pc/3151491?currency=USD&language=en_US')
for i in range(2, 9):
    gather('https://www.amazon.com/gp/bestsellers/pc/3151491?pg=' + str(i))

all['data'] = data
with open(json_file, 'w') as fp:
    json.dump(all, fp)

# br.close()

br.close()