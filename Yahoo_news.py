import requests
import urllib
import win32com.client

from bs4 import BeautifulSoup
from urllib.request import urlopen
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

category = "world"
title_num = 2
cmt_num = 2

browser = webdriver.Chrome('C:\\Users\\GunHee\\PycharmProjects\\data\\chromedriver')


excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")

ws.Cells(1,1).Value = "카테고리"
ws.Cells(1,2).Value = "기사 제목"

pages = 1

for i in range(1,1300,10):
    req = requests.get('https://news.yahoo.co.jp/list/?c='+ category +'&p='+str(i),
                     stream=True, headers={'User-agent': 'Mozilla/5.0'})
    html = req.text
    soup = BeautifulSoup(html, 'html.parser')

    test = soup.findAll("div", {"class" : "newsFeed_item_title"})
    print(test)
    for test2 in test :
        print('기사 제목 : '+test2.get_text())
        ws.Cells(title_num, 1).Value = "IT"
        ws.Cells(title_num, 2).Value = test2.get_text()
        # title_num = title_num + 1
        browser.get(test2.get('href'))
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        try:
            element = WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "news-comment-plguin-iframe"))
            )
            ws.Columns.AutoFit()
        except TimeoutException:
            print('기사없음')
        finally:
            print()

print("크롤링 끝")