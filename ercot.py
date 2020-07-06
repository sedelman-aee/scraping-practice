import requests
import csv
from bs4 import BeautifulSoup

#Get page
url = 'http://www.ercot.com/calendar/2020/1/7/189613-RMS'
page = requests.get(url)
soup = BeautifulSoup(page.text, 'html.parser')

#Create a file to write to
f = csv.writer(open('ercot-rms-docs.csv','w'))
f.writerow(['Name','Link','Posted On','Meeting','Type','Size'])

meeting = soup.find(class_ = 'pointer').contents[0]

section = soup.find(class_ = 'docsList')
list = section.find_all('li')

for item in list:

    doc = item.find('a')
    name = doc.contents[0]
    link = 'http://www.ercot.com' + doc.get('href')

    props = item.find(class_ = 'docProps').contents[0]
    props = props[1:-1]
    arr = props.split('–')
    posted = arr[0]
    type = arr[1]
    size = arr[2]

    f.writerow([name, link, posted[:-1], meeting, type[1:-1], size[1:]])
