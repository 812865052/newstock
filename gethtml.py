# -*- coding: utf-8 -*-
#
#Created on 2018年3月8日
#
#@author: sparrowzeng
#
import requests
import traceback
from bs4 import BeautifulSoup
import re
import xlwt
import xlrd

office_proxy = {
    "http": "http://proxy.tencent.com:8080",
    "https": "https://web-proxy.oa.com:8080"
}

dev_proxy = {
    "http": "http://dev-proxy.oa.com:8080",
    "https": "https://dev-proxy.oa.com:8080"
}

no_proxy = {
    "http": "",
    "https": ""
}
cookie = ''
WebBrowser = {
    "cookie": cookie,
    "proxy": no_proxy
}

class GetHtml():
    def __init__(self, webinfo):
        self.cookie = webinfo['cookie']
        self.proxy = webinfo['proxy']

    def getHtml(self, url):
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.119 Safari/537.36',
            'Cookie': self.cookie,
            'Connection': 'keep-alive'
        }
        self.url = url
        try:
            webdata = requests.get(url, headers=headers, timeout =20, proxies = self.proxy)
            self.html = webdata.text
            return self.html
        except:
            traceback.print_exc()
            print "There is something wrong with URL"
            return None

    def getSoup(self,url):
        html = self.getHtml(url)
        Soup = BeautifulSoup(html,'lxml')
        return Soup

if __name__ == '__main__':
    try:
        # 创建excel文件
        filename = xlwt.Workbook()
        # 给工作表命名，test
        sheet = filename.add_sheet("test")
        # 写入内容，第4行第3列写入‘张三丰’
        hello = u'张三丰'
        sheet.write(3, 2, hello)
        # 指定存储路径，如果当前路径存在同名文件，会覆盖掉同名文件
        path = "D:/test1.xls"
        filename.save(path)

    except Exception, e:
        print(str(e))

    # 打开excel文件
    date = xlrd.open_workbook(path)
    # 根据工作表名称，找到指定工作表  by_index(0)找到第N个工作表
    sheet = date.sheet_by_name('test')
    # 读取第四行第三列内容，cell_value读取单元格内容,指定编码
    value = sheet.cell_value(3, 2).encode('utf-8')
    print(value)

    url = "http://www.aastocks.com/sc/ipo/companysummary.aspx?symbol=08512&view=1"
    url2 = "http://www.aastocks.com/sc/ipo/IPOInfor.aspx?view=1&symbol=02048"
    url3 = "http://www.aastocks.com/sc/ipo/ListedIPO.aspx?iid=ALL&orderby=DA&value=DESC&index=2"
    # cookie = 'realLocationId=24; userFidLocationId=24; ip_ck=4cOJ5vr/j7QuNjQ4MTA4LjE1MjA1MDg3MDc%3D; listSubcateId=57; gr_user_id=50e8a30c-db29-4350-b900-beae64c4b153; Hm_lvt_ae5edc2bc4fc71370807f6187f0a2dd0=1520599840; z_day=iea1618%3D3%26izol100076%3D1%26ixgo20%3D1%26rdetail%3D7; z_pro_city=s_provice%3Dguangdong%26s_city%3Dshenzhen; userProvinceId=30; userCityId=348; userCountyId=0; userLocationId=24; record_number=1; lv=1521025082; vn=12; visited_subcateProId=57-1189144%2C1205708%2C1206990%2C1164497%2C1151984'
    browser = GetHtml(WebBrowser)
    Html = browser.getHtml(url3)
    Soup = browser.getSoup(url3)
    Html_nospace = Html.replace('\t', '').replace(' ', '').replace('\r\n', '').replace('&nbsp;', '')
    # Html_nospace = '<tdclass="defaulttitle">每手股数</td><tdstyle="padding-left:3px;">10000</td>'
    print Html_nospace
    # text = u'sss'
    # print type(text)
    valuereg = r'<tdclass=\"defaulttitle\">' + r'(.*?)?' + r'</td><tdstyle=\"padding-left:3px;\">(.*?)?</td>'
    valuereg2 = r'<tdclass=\"subtitle\">' + r'(.*?)?' + r'</td><tdclass=\"subtitle2\">(.*?)?</td>'
    valuereg3 = r'<tdclass=\"subtitle\">' + r'(.{,100}?)?' + r'</td><tdclass=\"subtitle3\">(.*?)?</td>'
    valuereg4 = r'>' + r'(.*?)?' + r'</a>\(<.*' + r'>' + r'(.*?)?' + r'</a>\(<'
    valuereg5 = r'<aclass=\"rlink\"href=\"/sc/stocks/quote/detail-quote.aspx\?symbol=[0-9]{5}\">' + r'([0-9]{5})?' + r'</a>'
    valuereg6 = r'<tdclass=\"rft\">' + r'(.*?)?' + r'</td><tdclass=\"rft\"><aclass=\"rlink\"href=\".*\">' + r'(.*?)?'
    valuereg7 = r'</a></td><tdclass=\"rftgbcolor\"><aclass=\"rlink\"href=\".*\">' + r'(.*?)?'
    valuereg8 = r'</a></td><tdclass=\"rgt\"><aclass=\"rlink\"href=\".*\">' + r'(.*?)?'
    valuereg9 = r'</a></td><tdclass=\"rgt\">' + r'(.*?)?' + r'</td><tdclass=\"rgt\">' + r'(.*?)?' + r'</td><tdclass=\"rgt\">' + r'(.*?)?'
    valuereg10 = r'</td><tdclass=\"rgt\">' + r'(.*?)?' + r'</td><tdclass=\"rgt\">' + r'(.*?)?' + r'</td><tdclass=\"rgt\">' + r'(.*?)?' + r'</td><tdclass=\"rgt\">' + r'(.*?)?' + r'</td><tdclass=\"rgt\">'
    tes = r'<tdclass=\"rft\">' \
        + r'(.*?)?' + r'</td>' + r'<tdclass=\"rft\"><aclass=\"rlink\"href=\"/sc/stocks/quote/detail-quote.aspx\?symbol=[0-9]{5}\">' \
        + r'(.*?)?' + r'</a></td>' + r'<tdclass=\"rftgbcolor\"><aclass=\"rlink\"href=\"CompanySummary.aspx\?view=1&Symbol=[0-9]{5}\">' \
        + r'(.*?)?' + r'</a></td>' + r'<tdclass=\"rgt\"><aclass=\"rlink\"href=\"IndustryComparison.aspx\?view=1&Symbol=[0-9]{5}\">' \
        + r'(.*?)?' + r'</a></td>' + r'<tdclass=\"rgt\">' \
        + r'(.*?)?' + r'</td>' + r'<tdclass=\"rgt\">' \
        + r'(.*?)?' + r'</td>' + r'<tdclass=\"rgt\">' \
        + r'(.*?)?' + r'</td>' + r'<tdclass=\"rgt\">' \
        + r'(.*?)?' + r'</td>' + r'<tdclass=\"rgt\"><spanclass=\"[a-z]{3}bold\">' \
        + r'(.*?)?' + r'</span></td>' + r'<tdclass=\"rgt\">' \
        + r'(.*?)?' + r'</td>' + r'<tdclass=\"rgt\"><spanclass=\"[a-z]{3}bold\">' \
        + r'(.*?)?' + r'</span></td>'
    # n = 0
    # for m in re.findall(tes, Html_nospace):
    #     n = n + 1
    #     for s in m:
    #         print s
    #         print n


    col = 1
    while(col<12):
        valuereg6 = r'<tdclass=\".*colorcol' + r'%sR{0,1}' %col + r'\"nowrap>(.*?)?</td>'
        valuetag_re = re.compile(valuereg6)
        for m in re.findall(valuetag_re, Html_nospace):
            print m
        col = col + 1


    # s = r'<tdclass=\"defaulttitle\">(.*?)?</td><tdstyle=\"padding-left:3px;\">(.*?)?</td>'
    # valuetag_re = re.compile(valuereg5)
    # for m in re.findall(valuetag_re, Html_nospace):
    #     # print type(Html_nospace)
    #     # print type(m[0])
    #     # print type(u'保荐人')
    #     print m
        # if(m[0]!= u'保荐人'):
        #     pass
        #     # print '1'
        #     # print m[1]
        # else:
        #     # print m[0]
        #     # print m[1]
        #     temp = m[1]
        #     # print temp
        #     valuetag_re = re.compile(valuereg4)
        #     for w in re.findall(valuetag_re, temp):
        #         print w[0],w[1]
    # for s in Soup.find_all(re.compile("^td")):
    #     print s


    # for s in Soup.find_all("td", class_="subtitle3"):
    #     print s
