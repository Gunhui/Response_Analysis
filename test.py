from urllib.request import urlopen
from urllib.request import HTTPError
from bs4 import BeautifulSoup
import json

try:
    html = urlopen("https://news.yahoo.co.jp/list/?c=world&p=1")
    bsobj = BeautifulSoup(html, "html.parser", from_encoding='utf-8')

except HTTPError as e:
    print(e)

array = [None]*6

for i in range(6):
    array[i] = dict([
        ("Category", None),
        ("Comment", None),
        ("img_src", None),
        ("img_href", None),
    ])
i = 0

for a in bsobj.select("dl.title > dt"):
    print(a.get_text())
# print(bsobj.find_all("dl.title > dt"))



# print(bsobj.find_all("a", {"class" : "newsFeed_item_link "}))

# print(bsobj.select('li.newsFeed_item'))
# print(bsobj.select("a"))
# print(bsobj.select("div > ul > li > a > div > div"))
# Category 크롤링
# for a in bsobj.select("a"):
#     href = a.attrs['href']
#     text = a.string
#     print(text, ">>", href)
try:
    for headline in bsobj.select("ul.newsFeed_list > li"):
        print(headline.get_text())
        print("Category : ", headline.get_text())
        print(i)
        array[i]['Category'] = headline.get_text()
        i+=1

# http((?!\").)*
except AttributeError as e:
    print(e)

i=0
# 주석 크롤링
try:
    for headline1 in bsobj.findAll('dd'):
        print(headline1.get_text())
        print(i)
        array[i]['Comment'] = headline1.get_text()
        i+=1

# http((?!\").)*
except AttributeError as e:
    print(e)

i=0
# img src & href 크롤링
test = bsobj.findAll('dl', attrs={"class" : "mtype_img"})

try:
    for headline1 in test:
        print(headline1.find('img').get('src'))
        print(headline1.find('a').get('href'))
        print('==========================')
        array[i]['img_src'] = headline1.find('img').get('src')
        array[i]['img_href'] = headline1.find('a').get('href')
        i+=1

# http((?!\").)*
except AttributeError as e:
    print(e)

print(array)
