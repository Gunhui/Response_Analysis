import tkinter
import requests
import urllib
import win32com.client

from bs4 import BeautifulSoup
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from googletrans import Translator
from urllib.request import urlopen
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

search_value = "science"
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

for i in range(1,100000):
    req = requests.get('https://news.yahoo.co.jp/list/?c='+ search_value +'&p='+str(i),
                 stream=True, headers={'User-agent': 'Mozilla/5.0'})
    html = req.text

    soup = BeautifulSoup(html, 'html.parser')

    test = soup.select('dl.title > dt')

    for test2 in test :
        print('기사 제목 : '+test2.text)
        ws.Cells(title_num, 1).Value = search_value
        ws.Cells(title_num, 2).Value = test2.text
        title_num += 1

print("크롤링 끝")