# -*-coding:gbk -*-

from bs4 import BeautifulSoup
import requesets
import xlrd
import xlwt
import xlutils.copy import copy
import time 
import winsound

#企查查网站爬虫类
class EnterpriseInfoSpider:
    def __init__(self):
        #文件相关
        self.excelPath = 'enterprise_data.xls'
        self.sheetName = 'details'
        self.workbook = None
        self.table = None
        self.beginRow = None
        
        #目录页
        self.catalogUrl = "http://www.qichacha.com/search_index"
        #详情页
        self.detailsUrl = "http://www.qichacha.com/company_getinfos"
        
        self.cookie = raw_input("请输入cookie：").decode("gbk").encode("utf-8")
        self.host = "www.qichacha.com"
        self.userAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.106 Safari/537.36"
        
        self.headers = {
            "cookie":self.cookie,
            "host":self.host,
            "user-agent":self.userAgent
        }
        
        #数据字段名17个
        self.fields = ['公司名称','电话号码','邮箱','统一社会信用代码','注册号','组织机构代码','经营状态','公司类型','成立日期','法定代表人','注册资本','营业期限','登记机关','发照日期','公司规模','所属行业','英文名','曾用名','企业地址','经营范围']
        
        #爬虫开始前的一些预处理
        def init(self):
            print("save...")
            book = xlwt.Workbook(encoding="utf-8",style_compression = 0)
            sheet = book.add_sheet("公司",cell_overwrite_ok=True)
            col = ('公司名称','电话号码','邮箱','统一社会信用代码','注册号','组织机构代码','经营状态','公司类型','成立日期','法定代表人','注册资本','营业期限','登记机关','发照日期','公司规模','所属行业','英文名','曾用名','企业地址','经营范围')
            for i in range(len(col)):
                sheet.write(0,i,col[i])
            for i in range(0,len(datalist)):
                print("第%d条数据"%i)
                data = datalist[i]
                for j in range(4):
                    sheet.write(i+1,j,data[j])
            book.save(savepath)
            
        def start(self):
            keyword = raw_input()