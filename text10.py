#coding = utf-8
import sqlite3

from bs4 import BeautifulSoup
import re
import urllib.request,urllib.error
import xlwt



def main():
    baseurl = "http://movie.douban.com/top250?start="
    #1.爬取网页
    datalist = getData(baseurl)
    #2.解析数据
    savepath = ".\\豆瓣电影Top250.xls"
    #dbpath = "movie.db"
    #3.保存数据
    saveData(datalist,savepath)
   # saveData2DB(datalist,dbpath)
    #askURL("http://movie.douban.com/top250?start=")
 #抓取网页

 #影片详情链接的规则
findLink =re.compile(r'<a href="(.*?)">')      #创建正则表达式对象，表示规则（字符串的模式）
#影片的图片
findImgSrc = re.compile(r'<img.*src="(.*?)"',re.S)   #re.S让换行符包括进去
#影片的片名
findTitle = re.compile(r'<span class="title">(.*)</span>')
#影片的评分
findRating = re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')
#评价人数
findJudge = re.compile(r'<span>(\d*)人评价</span>')
#具体概括
findInq = re.compile(r'<span class="inq">(.*)</span>')
#找到影片的相关内容
findBd = re.compile(r'<p class="">(.*?)</p>',re.S)



def getData(baseurl):
    datalist = []
    for i in range(0,10):  #调用获取页面信息的函数，10次
        url = baseurl +str(i*25)
        html = askURL(url)  #保存获取到的网页源码
        #2.逐一解析数据
        soup = BeautifulSoup(html,"html.parser")
        for item in soup.find_all('div',class_= "item"):#查找符合要求的字符串，形成列表
         #print(item)
            data = []  #保存一部电影的所有信息
            item = str(item)


#获取到影片详情的链接
            link = re.findall(findLink,item)[0]   #通过正则表达式查找指定的字符串
            data.append(link)
            imgSrc = re.findall(findImgSrc,item)[0]
            data.append(imgSrc)                   #添加图片
            titles = re.findall(findTitle,item)   #片名可能只有一个中文名，没有外国名
            if (len(titles) == 2):
                ctitle =titles[0]
                data.append(ctitle)              #添加中文名
                otitle= titles[1].replace("/","")  #去掉无关的符号
                data.append(otitle)                 #添加外国名
            else:
                data.append(titles[0])
                data.append(' ')                #外国名字留空

            rating = re.findall(findRating,item)[0]
            data.append(rating)

            judgeNum = re.findall(findJudge,item)[0]
            data.append(judgeNum)

            inq = re.findall(findInq,item)
            if len(inq)!= 0:
                inq = inq[0].replace("。","")    #去掉句号
                data.append(inq)
            else:
                data.append(" ")

            bd = re.findall(findBd,item)[0]
            bd = re.sub('br(\s+)?/>(\s+)?'," ",bd)  #去掉<br/>
            bd = re.sub('/'," ",bd)
            data.append(bd.strip())  #去掉前后的空格

            datalist.append(data)    #把处理好的数据存到一张表
    print(datalist)

    return datalist

#3.保存数据
def saveData(datalist,savepath):
    print("save ...")

    book = xlwt.Workbook(encoding="utf-8",style_compression=0)   #创建workbook对象
    sheet = book.add_sheet('豆瓣电影Top250',cell_overwrite_ok=True)  #创建工作表
    col = ("电影详情链接","图片链接","影片中文名","影片外国名","评分","评价数","概识","相关信息")
    for i in range(0,8):
        sheet.write(0,i,col[i])  #列名
    for i in range(0,250):
        print("第%d条"%i)
        data = datalist[i]
        for j in range(0,8):
            sheet.write(i+1,j,data[j])  #数据

    book.save(savepath)        #保存


def saveData2DB(datalist,dbpath):
    print("...")

def init_db(dbpath):
    sql = '''
       create table movie250
       (
       id integer primary key autoincrement,
       info_link text,
       pic_link text,
       cname varchar,
       ename varchar,
       score numeric,
       rated numeric,
       introduction text,
       info text
       )
    ''' #创建表格
    conn = sqlite3.connect(dbpath)
    cursor = conn.cursor()
    cursor.execute(sql)
    conn.commit()
    conn.close()

#得到制定的一个URL的网页内容
def askURL(url):
    head = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.190 Safari/537.36"}
    #用户代理表示告诉豆瓣服务器，我们是什么类型的机器，浏览器（本质上干是告诉浏览器，我们可以接受什么水平的文件内容）
    request = urllib.request.Request(url,headers=head)
    html = ""
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8")
       # print(html)
    except urllib.error.URLError as e:
        if hasattr(e,"code"):
            print(e.code)
        if hasattr(e,"reason"):
            print(e.reason)
    return html






if __name__ == "__main__":
    main()
    #init_db("movie.db")
    print('爬取完毕。')