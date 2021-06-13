import math
import os
import sys
import time
import urllib.request
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver

query_txt = 'Gmarket'
query_url = 'http://corners.gmarket.co.kr/Bestsellers'

cnt = int(input('1.크롤링 할 건수는 몇건입니까?: '))
real_cnt = math.ceil(cnt / 20)

f_dir = input('2.파일이 저장될 경로만 쓰세요;(예:c:\\temp\\): ')
print("\n")

now = time.localtime()
s = '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

img_dir = f_dir + s + '-' + query_txt + "\\images"

os.makedirs(img_dir)

os.chdir(f_dir + s + '-' + query_txt)

f_name = f_dir + s + '-' + query_txt + '\\' + s + '-' + query_txt + '.txt'
fc_name = f_dir + s + '-' + query_txt + '\\' + s + '-' + query_txt + '.csv'
fx_name = f_dir + s + '-' + query_txt + '\\' + s + '-' + query_txt + '.xls'

s_time = time.time()

path = "c:/temp/chromedriver_240/chromedriver.exe"
driver = webdriver.Chrome(path)

driver.get(query_url)
time.sleep(1)


def scroll_down(driver):
    driver.execute_script("window.scrollTo(0, window.scrollY + 1000);")
    time.sleep(1)

i = 1
while (i <= 20):
    scroll_down(driver)
    i += 1

rank2 = []
title2 = []
cost_price2 = []
sale_price2 = []
discount_rate2 = []

html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

count = 1

sale_result = soup.select('div.best-list')
slist = sale_result[1].select('ul > li')

img_file_no = 0

for li in slist:

    os.chdir(img_dir)

    try:
        photo = li.find('div', 'thumb').find('img')['src']
    except AttributeError:
        continue

    img_file_no += 1

    urllib.request.urlretrieve(photo, str(img_file_no) + '.jpg')
    time.sleep(2)

    if img_file_no > cnt:
        break

for li in slist:

    f = open(f_name, 'a', encoding='UTF-8')
    f.write("-----------------------------------------------------" + "\n")

    print("\n")
    print("-" * 70)
    sid = '#no' + str(count)
    try:
        rank = li.select(sid)[0].get_text()
    except AttributeError:
        rank = ''
        print('1.판매순위:', rank.replace("\n", ""))
    else:
        print("1.판매순위:", rank)

        f.write('1.판매순위:' + rank + "\n")

    try:
        title = li.select('a.itemname')[0].get_text()
    except AttributeError:
        title = ''
        print(title)
        f.write('2.제품소개:' + title + "\n")
    else:
        print("2.제품소개:", title.replace("\n", ""))
        f.write('2.제품소개:' + title + "\n")

        cost_price = li.find('div', class_='item_price').find('div', 'o-price').get_text()
        print("3.원래가격:", cost_price.replace("\n", ""))
        f.write('3.원래가격:' + cost_price + "\n")

        sale_price = li.find('div', class_='item_price').find('div', 's-price').find('strong').get_text()
        print("4.판매가격:", sale_price.replace("\n", ""))
        f.write('4.판매가격:' + sale_price + "\n")

        try:
            discount_rate = li.find('div', class_='item_price').find('div', 's-price').find('em').get_text()
        except AttributeError:
            discount_rate = '0%'
        print("5.할인율:", discount_rate.replace("\n", ""))
        f.write('5.할인율:' + discount_rate + "\n")

        rank2.append(rank)
        title2.append(title.replace("\n", ""))
        cost_price2.append(cost_price.replace("\n", ""))
        sale_price2.append(sale_price.replace("\n", ""))
        discount_rate2.append(discount_rate.replace("\n", ""))

        if count == cnt:
            break

        count += 1

        time.sleep(0.5)
g_best_seller = pd.DataFrame()

g_best_seller['판매순위'] = rank2
g_best_seller['제품소개'] = pd.Series(title2)
g_best_seller['원래가격'] = pd.Series(cost_price2)
g_best_seller['판매가격'] = pd.Series(sale_price2)
g_best_seller['할인율'] = pd.Series(discount_rate2)

g_best_seller.to_csv(fc_name, encoding="utf-8-sig", index=True)

g_best_seller.to_excel(fx_name, index=True)

e_time = time.time()
t_time = e_time - s_time

orig_stdout = sys.stdout
f = open(f_name, 'a', encoding='UTF-8')
sys.stdout = f

sys.stdout = orig_stdout
f.close()

print("\n")
print("1.파일 저장 완료: txt 파일명 : %s " % f_name)
print("2.파일 저장 완료: csv 파일명 : %s " % fc_name)
print("3.파일 저장 완료: xls 파일명 : %s " % fx_name)

import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fx_name)
sheet = wb.ActiveSheet
sheet.Columns(3).ColumnWidth = 30
row_cnt = cnt + 1
sheet.Rows("2:%s" % row_cnt).RowHeight = 120

ws = wb.Sheets("Sheet1")
col_name2 = []
file_name2 = []

for a in range(2, cnt + 2):
    col_name = 'C' + str(a)
    col_name2.append(col_name)

for b in range(1, cnt + 1):
    file_name = img_dir + '\\' + str(b) + '.jpg'
    file_name2.append(file_name)

for i in range(0, cnt):
    rng = ws.Range(col_name2[i])
    image = ws.Shapes.AddPicture(file_name2[i], False, True, rng.Left, rng.Top, 130, 100)
    excel.Visible = True
    excel.ActiveWorkbook.Save()

driver.close()
