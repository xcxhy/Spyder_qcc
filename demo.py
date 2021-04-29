#encoding:utf-8
import sqlite3
from bs4 import BeautifulSoup
import re
import urllib.request,urllib.error,urllib.parse
import xlwt
import os
import pandas as pd
import socket
socket.setdefaulttimeout(30)
# import importlib,sys
# importlib.reload(sys)
# import sys
# import io
# sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf8')

def main():
    baseurl = "https://www.qcc.com/web/search?key="
    #1.爬取网页
    datalist =  getData(baseurl)
    #2.解析数据
    
    #3.保存数据
    savepath = ".\\company.xls"
    saveData(datalist,savepath)

find_address = re.compile(r'<spanclass="f"data-v-681cda67="">地址：<spanclass="val"data-v-681cda67="">(.*)</span>')
find_date = re.compile(r'</span><spanclass="f"data-v-681cda67="">成立日期：<spanclass="val"data-v-681cda67="">(.*)</span>')
find_money = re.compile(r'<spanclass="f"data-v-681cda67="">注册资本：<spanclass="val"data-v-681cda67="">(.*)</span>')

def getData(baseurl):
    datalist = []
    data_names = pd.read_excel(r'C:\Users\xcxhy\Desktop\2.xlsx')
    data_names = data_names.iloc[:,1]
    data_names = data_names.values
    for i in range(len(data_names)):
        data_name = data_names[i]
        if (data_name[0] >'a' and data_name[0] < 'z') or (data_name[0] >'A' and data_name[0] <'Z'):
            continue
        url = baseurl + urllib.parse.quote(data_names[i])
        print(url)
        html = askURL(url) 
        #print(html)
        #保存获取到的网页源码
        #2.逐一解析数据
        soup = BeautifulSoup(html,"html.parser")
    
        #print(soup)
        try:
            for item in soup.find('table',"ntable ntable-list"):
                
                data = []
                item =str(item)
                item = item.replace("\n","")
                item = item.replace(" ","")
                print(item)
                name = data_names[i]
                data.append(name)
                address = re.findall(find_address,item)
                address = re.sub('</span(.*)',"",address)
                data.append(address)
                date = re.findall(find_date,item)
                date = re.sub('</span(.*)',"",date)
                data.append(date)
                money = re.findall(find_money,item)
                re.sub('</span(.*)',"",money)
                data.append(money)
                datalist.append(data)
        except TypeError as e:
            print("1111")
            pass
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
    
    head = {"cookie":cookie,"method": "GET","User-Agent":"Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN) AppleWebKit/523.15 (KHTML, like Gecko, Safari/419.3) Arora/0.3 (Change: 287 c9dfb30)"}
    proxy = urllib.request.ProxyHandler({"http":"222.73.130.111:8080"})
    opener = urllib.request.build_opener(proxy,urllib.request.HTTPHandler)
    urllib.request.install_opener(opener)

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
    cookie = input("请输入cookie:").encode("utf-8")
    main()
    
    print("爬取完毕")