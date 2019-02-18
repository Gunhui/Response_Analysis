import urllib
import win32com.client
import tkinter as tk
import requests
import matplotlib.pyplot as plt
import matplotlib.figure
import matplotlib.patches

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from pandas import DataFrame
from tkinter import *
from bs4 import BeautifulSoup
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from googletrans import Translator
from urllib.request import urlopen
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


main = Tk()
main.title("Comment Analyze")
main.geometry('640x480+100+100')
lbl = Label(main, text="Enter the Keyword")
lbl.grid(column=0, row=0)
lbl.config(font=("Courier", 20))
lbl.place(x=210, y=170)


txt = Entry(main, width=20)
txt.grid(column=1, row=0)
txt.place(x=260, y=270)
txtLb = Label(main, text="Text")
txtLb.grid(column=0, row=0)
txtLb.config(font=("Courier", 10))
txtLb.place(x=210, y=270)

num = Entry(main, width=10)
num.grid(column=1, row=0)
num.place(x=260, y=220)
numLb = Label(main, text="Page")
numLb.grid(column=0, row=0)
numLb.config(font=("Courier", 10))
numLb.place(x=210, y=220)



def clicked():
    search_value = txt.get()
    title_num = 2
    cmt_num = 2

    news_result = []
    news = {}
    news_comment = {}
    pos = 0
    neg = 0
    neu = 0

    analyser = SentimentIntensityAnalyzer()
    translator = Translator()

    browser = webdriver.Chrome('C:\\Users\\GunHee\\PycharmProjects\\data\\chromedriver')

    # excel = win32com.client.Dispatch("Excel.Application")
    # excel.Visible = True
    # wb = excel.Workbooks.Add()
    # ws = wb.Worksheets("Sheet1")

    # ws.Cells(1,1).Value = "기사 제목"
    # ws.Cells(1,2).Value = "기사 댓글"

    def print_sentiment_scores(sentence):
        division = str()
        snt = analyser.polarity_scores(sentence)
        print(str(snt['compound']))
        if(snt['compound']>0.0):
            division = "pos"
        elif(snt['compound']<0.0):
            division = "neg"
        else:
            division="neu"

        return division

    pages = 1

    for i in range(1, int(num.get()), 10):
        req = requests.get(
            'https://news.yahoo.co.jp/search/?fr=top_ga1_sa&p=' + search_value + '&oq=' + search_value + '&ei=UTF-8&b=' + str(i),
            stream=True, headers={'User-agent': 'Mozilla/5.0'})
        html = req.text
        soup = BeautifulSoup(html, 'html.parser')
        move_news_tags = soup.select('div > div > h2 > a')

        for move_news_tag in move_news_tags:
            print('기사 제목 : ' + move_news_tag.text)
            '''
           for excel
           ws.Cells(title_num, 1).Value = move_news_tag.text
           title_num = title_num + 1
          '''

            browser.get(move_news_tag.get('href'))
            browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            try:
                element = WebDriverWait(browser, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "news-comment-plguin-iframe"))
                )

                news_html = browser.page_source

                news_soup = BeautifulSoup(news_html, 'html.parser')
                cmt = news_soup.find('iframe', attrs={'class': 'news-comment-plguin-iframe'})

                cmt_req = requests.get(cmt.get('src'))

                cmt_html = cmt_req.text
                cmt_soup = BeautifulSoup(cmt_html, 'html.parser')
                comments = cmt_soup.findAll('span', attrs={'class': 'cmtBody'})

                for comment in comments:
                    print(cmt_num)
                    trans = translator.translate(comment.text, src="ja", dest="en")
                    print(' 댓글  ' + comment.text)

                    feel = print_sentiment_scores(trans.text)
                    news_comment[comment.text] = feel
                    # ws.Cells(cmt_num, 2).Value = comment.text + feel
                    cmt_num = cmt_num + 1
                    title_num = title_num + 1
                print()
                news[move_news_tag.text] = news_comment
                news_comment = {}

            except TimeoutException:
                print('기사없음')
            finally:
                print()

    news_result.append(news)

    analyze = Tk()
    analyze.title("Comment Analyze")
    analyze.geometry("1300x700")

    result = Label(analyze, text="Result")
    result.grid(column=0, row=0)
    result.config(font=("Courier", 20))
    result.place(x=580, y=30)

    graph = Label(analyze, text="Graph")
    graph.grid(column=0, row=0)
    graph.config(font=("Courier", 20))
    graph.place(x=580, y=300)

    frame = Frame(analyze)
    frame.place(x=70, y=80)

    listNodes = Listbox(frame, width=190, height=13, font=("Helvetica", 8))
    listNodes.pack(side="left", fill="y")

    scrollbar = Scrollbar(frame, orient="vertical")
    scrollbar.config(command=listNodes.yview)
    scrollbar.pack(side="right", fill="y")

    listNodes.config(yscrollcommand=scrollbar.set)

    tmp = "=============================================================================================================================================================================================="
    tmp1 = "Comment List"

    for i in range(len(news_result[0])):
        listNodes.insert(END, list(news_result[0].keys())[i])

        for j in range(len(list(news_result[0].values())[i])):
            listNodes.insert(END, tmp1)
            listNodes.insert(END, list(list(news_result[0].values())[i].keys())[j])
            if(list(list(news_result[0].values())[i].values())[j] == "pos"):
                pos+=1
            elif(list(list(news_result[0].values())[i].values())[j] == "neu"):
                neu+=1
            else:
                neg+=1
        listNodes.insert(END, tmp)

    tmp2 = Label(analyze)
    tmp2.grid(column=0, row=0)
    tmp2.place(x=360, y=300)

    Data1 = {'Country': ['Pos', 'Neg', 'Neu'],
             'Emotion': [pos, neg, neu]
             }

    df1 = DataFrame(Data1, columns=['Country', 'Emotion'])
    df1 = df1[['Country', 'Emotion']].groupby('Country').sum()

    figure1 = plt.Figure(figsize=(5, 3.5), dpi=100)
    ax1 = figure1.add_subplot(111)
    bar1 = FigureCanvasTkAgg(figure1, tmp2)
    bar1.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH)
    df1.plot(kind='bar', legend=True, ax=ax1, color=tuple(['r', 'b', 'g']))
    ax1.set_title('Response Analysis')

    analyze.mainloop()

btn = Button(main, text="Search", command=clicked)
btn.grid(column=2, row=0)
btn.place(x=300, y=305)

main.mainloop()