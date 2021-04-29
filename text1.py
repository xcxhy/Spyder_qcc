#encoding:utf-8
import sqlite3
from bs4 import BeautifulSoup
import re
import urllib.request,urllib.error,urllib.parse
import xlwt
import os
import pandas as pd
# import importlib,sys
# importlib.reload(sys)
# import sys
# import io
# sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf8')

def main():
    baseurl = "https://www.baidu.com"
    #1.爬取网页
    datalist =  getData(baseurl)
    #2.解析数据
    
    #3.保存数据
    savepath = ".\\company.xls"
    saveData(datalist,savepath)

find_address = re.compile(r'<span (.*) class="f">"地址："<span (.*) class="val">(.*)</span>')
find_date = re.compile(r'<span (.*) class="f">"成立日期："<span (.*) class="val">(.*)</span>')
find_money = re.compile(r'<span (.*) class="f">"注册资本："<span (.*) class="val">(.*)</span>')

def getData(baseurl):
    datalist = []
    data_names = pd.read_excel(r'C:\Users\xcxhy\Desktop\company.xlsx')
    data_names = data_names.iloc[:,1]
    data_names = data_names.values
    for i in range(len(data_names)):
        url = baseurl
        print(url)
        html = askURL(url) 
        print(html)
        #保存获取到的网页源码
        #2.逐一解析数据
        soup = BeautifulSoup(html,"html.parser")
        print(soup)
        for item in soup.find(text = re.compile(r'tr data-v-681cda67 class')):
            print(item)
            data = []
            item =str(item)
            name = data_names[i]
            data.append(name)
            address = re.findall(find_address,item)[0]
            data.append(address)
            date = re.findall(find_date,item)[0]
            data.append(date)
            money = re.findall(find_money,item)[0]
            data.append(money)
            datalist.append(data)
    return datalist
#保存数据
def saveData(datalist,savepath):
    print("save...")
    book = xlwt.Workbook(encoding="utf-8",style_compression = 0)
    sheet = book.add_sheet("公司",cell_overwrite_ok=True)
    col = ("公司名称","公司地址","注册时间","注册资金")
    for i in range(4):
        sheet.write(0,i,col[i])
    for i in range(0,len(datalist)):
        print("第%d条数据"%i)
        data = datalist[i]
        for j in range(4):
            sheet.write(i+1,j,data[j])
    book.save(savepath)

def askURL(url):
    head = {"method": "GET","User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.190 Safari/537.36"}
    

    request = urllib.request.Request(url,headers = head)
    html = ""
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8","ignore")
        #print(html)
    except urllib.error.URLError as e:
        if hasattr(e,"code"):
            print(e.code)
        if hasattr(e,"reason"):
            print(e.reason)
    return html

if __name__ == "__main__":
    main()
    print("爬取完毕")
