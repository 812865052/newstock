# -*- coding: utf-8 -*-
#
#Created on 2018年7月27日
#
#@author: feiye
#

import os,re
import gethtml
import xlwt,xlrd
from xlrd import open_workbook
from xlutils.copy import copy

def createtable(path,url):
    '''
    其实这里最好拆分成两个函数
    如果没有excel表格，就创建excel表格，如果有，则返回表格内容，行列数等
    :param path: 表格路径
    :param url: 需要解析url的路径
    :return: Html_nospace（url对应网页的内容）, table（创建excel的对象）, excel, rows, cols
    '''

    #爬网页内容
    browser = gethtml.GetHtml(gethtml.WebBrowser)
    Html = browser.getHtml(url)
    Soup = browser.getSoup(url)
    Html_nospace = Html.replace('\t', '').replace(' ', '').replace('\r\n', '').replace('&nbsp;', '')

    #如果文件不存在，则创建文件先
    if os.path.exists(path) != True:
        try:
            # 创建excel文件
            filename = xlwt.Workbook()
            # 给工作表命名，test
            sheet = filename.add_sheet("test")
            sheet._cell_overwrite_ok = True
            filename.save(path)
        except Exception, e:
            print(str(e))

    #采用追加方式写excel文件
    rexcel = open_workbook(path)  # 用wlrd提供的方法读取一个excel文件
    rows = rexcel.sheets()[0].nrows  # 用wlrd提供的方法获得现在已有的行数
    cols = rexcel.sheets()[0].ncols  # 用wlrd提供的方法获得现在已有的列数
    excel = copy(rexcel)  # 用xlutils提供的copy方法将xlrd的对象转化为xlwt的对象
    table = excel.get_sheet(0)  # 用xlwt对象的方法获得要操作的sheet

    return Html_nospace, table, excel, rows, cols


def readcompanynum(path,url):
    '''
    读取excel表格的第一列的股票代码列表
    :param path:
    :param url:
    :return:
    '''

    #如果文件不存在，则创建文件先
    if os.path.exists(path) != True:
        try:
            # 创建excel文件
            filename = xlwt.Workbook()
            # 给工作表命名，test
            sheet = filename.add_sheet("test")
            sheet._cell_overwrite_ok = True
            filename.save(path)
        except Exception, e:
            print(str(e))

    # 打开excel文件
    date = xlrd.open_workbook(path)
    # 根据工作表名称，找到指定工作表  by_index(0)找到第N个工作表
    sheet = date.sheet_by_name('test')
    rows = sheet.nrows  # 获取行数
    cols = sheet.ncols  # 获取列数
    # 读取第四行第三列内容，cell_value读取单元格内容,指定编码
    tmp = 1
    companynum = []
    while(tmp<rows):
        # print sheet.cell_value(tmp, 1).encode('utf-8')
        companynum.append(sheet.cell_value(tmp, 1).encode('utf-8'))
        tmp = tmp + 1

    print companynum
    return companynum



def createcompanytable(path,url,rowlist):
    '''
    填充所有表头，保值每列的表头都填正确
    :param path: excel的路径
    :param url: 获取对应url网页内容，然后解析
    :param rowlist: 表头list
    :return:
    '''

    Html_nospace, table, excel, rows, cols= createtable(path,url)

    # 从网站抓取特定字段做表头
    col = 1
    while (col < 12):
        valuereg6 = r'<tdclass=\".*colorcol' + r'%sR{0,1}' % col + r'\"nowrap>(.*?)?</td>' #%s 就是后面的%col的内容
        valuetag_re = re.compile(valuereg6)
        print
        for m in re.findall(valuetag_re, Html_nospace):
            if m.find(u'首日表現') != -1:
                # print m[:m.find('<sup')]
                m = u'首日表現'
            if m.find(u'现价') != -1:
                # print m[:m.find('<sup')]
                m = u'现价'
            table.write(0, col - 1, m)
        col = col + 1

    #下面的是补充从网站无法抓取到的表头，或者是网站没有的，需要自己手动加上的字段，他们的位置
    #就是挨着从网页内容读取的表头的后面
    temp = 0
    while(temp<len(rowlist)):
        table.write(0, col+temp-1, rowlist[temp])
        temp = temp + 1

    excel.save(path)  # xlwt对象的保存方法，这时便覆盖掉了原来的excel



def fillcompanytable(path,url):
    Html_nospace, table, excel, rows, cols= createtable(path, url)
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
    row = rows
    col_num = 1
    for m in re.findall(tes, Html_nospace):
        col = 0
        # print m[1] #从整个股票简易信息中提取出股票代码
        exist = search(m[1],path,col_num)
        if(exist!='not found'):
            pass
        else:
            for s in m:
                if(s=='N/A'):
                    s = '0'
                table.write(row, col, s)
                col = col + 1
        row = row + 1
    excel.save(path)


def detailcompanytable(path,url):
    '''
    每个股票代码对应的全部详细信息
    :param path:
    :param url:
    :return:
    '''
    Html_nospace, table, excel, rows, cols= createtable(path, url)
    dict = {}

    #筛选上半部分基本信息
    valuereg = r'<tdclass=\"subtitle\">' + r'(.*?)?' + r'</td><tdclass=\"subtitle2\">(.*?)?</td>'
    valuetag_re = re.compile(valuereg)
    for m in re.findall(valuetag_re, Html_nospace):
        first = m[0]
        first2 = m[0]
        if m[0].find(u'全球发售股数') != -1:
            first = u'全球发售股数'
        elif m[0].find(u'香港/公开发售股数') != -1:
            first = u'香港/公开发售股数'
        elif m[0].find(u'国际配售股数') != -1:
            first = u'国际配售股数'
        elif m[0].find(u'招股价') != -1:
            first = u'招股价下限'
            first2 = u'招股价上限'
        # print first
        # print m[1]
        if(first!=u'招股价下限'):
            dict[first] = m[1]
        else:
            # print m[1].split('-')
            # print type(m[1].split('-'))
            if(len(m[1].split('-'))!=1):
                print float(m[1].split('-')[0])
                dict[first] = float(m[1].split('-')[0])
                dict[first2] = float(m[1].split('-')[1])
            else:
                if(m[1].split('-')[0]=='N/A'):
                    dict[first] = 0
                    dict[first2] = 0
                else:
                    # print float(m[1].split('-')[0])
                    dict[first] = float(m[1].split('-')[0])
                    dict[first2] = float(m[1].split('-')[0])

    # 筛选下半部分基本信息
    valuereg = r'<tdclass=\"subtitle\">' + r'(.{,100}?)?' + r'</td><tdclass=\"subtitle3\">(.*?)?</td>'
    valuereg4 = r'<aclass=\'.*?\'href=\'.*?\'>' + r'(.*?)?' + r'</a>'  # 经典，非贪婪匹配的最佳例子，注意输出对比
    valuetag_re = re.compile(valuereg)
    flag_bjr = False
    baojianren = []
    for m in re.findall(valuetag_re, Html_nospace):
        second = m[1]
        if (m[0] != u'保荐人'):
            pass
            # print m[0], m[1]
        else:
            if (flag_bjr == False):
                temp = m[1].replace(u'(', '').replace(u')', '').replace(u'、', '')
                # print temp.find('\'>')
                valuetag_re1 = re.compile(valuereg4)
                for w in re.findall(valuetag_re1, temp):
                    if (w != u'相关往绩'):
                        baojianren.append(w)
                flag_bjr = True
            else:
                pass
            second = baojianren
        dict[m[0]] = second
    # for s in baojianren:
    #     print s
    # for s in dict:
    #     print s,dict[s]

    return dict

def searchfillcompany(compannynum,path):
    '''
    根据compannynum中的股票代码的list，匹配每个股票代码，然后去特定网址中找每个股票代码的详细信息，然后填到excel
    :param compannynum:
    :param path: excel路径
    :return:
    '''
    Html_nospace, table, excel, rows, cols= createtable(path, url)
    # 打开excel文件
    date = xlrd.open_workbook(path)
    # 根据工作表名称，找到指定工作表  by_index(0)找到第N个工作表
    sheet = date.sheet_by_name('test')
    rows = sheet.nrows  # 获取行数
    cols = sheet.ncols  # 获取列数

    # print sheet.cell_value(1,1).encode('utf-8')

    row_title = 0
    col_num = 1
    for num in compannynum:
        num_row = search(num,path,col_num)
        url_detail = "http://www.aastocks.com/sc/ipo/IPOInfor.aspx?view=1&symbol=" + num
        dict = detailcompanytable(path,url_detail)
        for s in dict:
            col = searchtitle(s, path, row_title)
            if(col!='not found'):
                table.write(num_row, col, dict[s])

    excel.save(path)


def search(num,path,col):
    # 打开excel文件
    date = xlrd.open_workbook(path)
    # 根据工作表名称，找到指定工作表  by_index(0)找到第N个工作表
    sheet = date.sheet_by_name('test')
    rows = sheet.nrows  # 获取行数
    cols = sheet.ncols  # 获取列数

    row = 1
    while(row<rows):
        if(num==sheet.cell_value(row,col).encode('utf-8')):
            # print row,sheet.cell_value(row,col).encode('utf-8')
            return row
        row = row + 1

    return 'not found'


def searchtitle(title,path,row):
    '''
    返回title在excel中所属表头的列数
    :param title: 待查表头
    :param path: excel路径
    :param row: 在哪一行进行查找（表头是0）
    :return:
    '''
    # 打开excel文件
    date = xlrd.open_workbook(path)
    # 根据工作表名称，找到指定工作表  by_index(0)找到第N个工作表
    sheet = date.sheet_by_name('test')
    rows = sheet.nrows  # 获取行数
    cols = sheet.ncols  # 获取列数

    col = 0
    while(col<cols):
        if(title==sheet.cell_value(row,col)):
            # print col,sheet.cell_value(row,col).encode('utf-8')
            return col
        col = col + 1

    return 'not found'


def computeexcelnum(path,descol,sourcecol):
    '''
    计算20手会平均中多少股
    :param path: excel路径
    :param descol: 目标列表头
    :param sourcecol: 计算列列表的表头
    :return:
    '''
    row_title = 0
    dict = {}
    for s in sourcecol:
        dict[s] = searchtitle(s,path,row_title)
    dict[descol] = searchtitle(descol, path, row_title)

    Html_nospace, table, excel, rows, cols = createtable(path, url)
    # 打开excel文件
    date = xlrd.open_workbook(path)
    # 根据工作表名称，找到指定工作表  by_index(0)找到第N个工作表
    sheet = date.sheet_by_name('test')
    rows = sheet.nrows  # 获取行数
    cols = sheet.ncols  # 获取列数

    # print sheet.cell_value(1,1).encode('utf-8')

    row = 1
    while(row<rows):
        para = len(sourcecol)
        # print sourcecol[0],sheet.cell_value(row,dict[sourcecol[0]])
        # print sourcecol[1], sheet.cell_value(row, dict[sourcecol[1]])
        if(dict[sourcecol[1]]=='N/A'):
            dict[sourcecol[1]] = '0%'
        value = 20*float(sheet.cell_value(row,dict[sourcecol[0]]))*(float(sheet.cell_value(row,dict[sourcecol[1]]).strip('%'))/100)
        table.write(row,dict[descol],value)
        row = row + 1

    excel.save(path)


def computeexcelmoney(path,descol,sourcecol):
    '''
    计算招股一手金额
    :param path: excel路径
    :param descol: 目标列
    :param sourcecol: 多个计算列的表头列表
    :return:
    '''
    row_title = 0
    dict = {}
    for s in sourcecol:
        dict[s] = searchtitle(s,path,row_title)
    dict[descol] = searchtitle(descol, path, row_title)

    Html_nospace, table, excel, rows, cols = createtable(path, url)
    # 打开excel文件
    date = xlrd.open_workbook(path)
    # 根据工作表名称，找到指定工作表  by_index(0)找到第N个工作表
    sheet = date.sheet_by_name('test')
    rows = sheet.nrows  # 获取行数
    cols = sheet.ncols  # 获取列数

    # print sheet.cell_value(1,1).encode('utf-8')

    row = 1
    while(row<rows):
        para = len(sourcecol)
        value = float(sheet.cell_value(row,dict[sourcecol[0]]))*float(sheet.cell_value(row,dict[sourcecol[1]]))
        table.write(row,dict[descol],value)
        row = row + 1

    excel.save(path)

if __name__ == '__main__':
    # path = 'D:/test3.xls'
    path = '/users/yefei/documents/newstock.xls'
    rowlist = [u'每手股数', u'全球发售股数', u'招股价下限', u'招股价上限', u'香港/公开发售股数', u'国际配售股数', u'市值', u'保荐人',u'招股一手金额',u'是否有旧股',u'基础投资人占比', \
               u'20账户一手申请可中股数',u'暗盘收盘价',u'开盘固定时点价',u'盈利']

    url = "http://www.aastocks.com/sc/ipo/ListedIPO.aspx?iid=ALL&orderby=DA&value=DESC&index=2"
    # url_detail = "http://www.aastocks.com/sc/ipo/IPOInfor.aspx?view=1&symbol=02048"


    pagenum = 2
    tmp = 1
    createcompanytable(path, url, rowlist)

    #读取不同页面所包含的股票代码列表
    while(tmp<pagenum):
        url = "http://www.aastocks.com/sc/ipo/ListedIPO.aspx?iid=ALL&orderby=DA&value=DESC&index=" + str(tmp)
        fillcompanytable(path, url)
        # search('02003',path,1)
        # searchtitle('每手股数',path,0)
        companynum = readcompanynum(path,url)
        searchfillcompany(companynum,path)
        tmp = tmp + 1

    descol = u'20账户一手申请可中股数'
    sourcecol = [u'每手股数', u'中签率(%)']
    descol1 = u'招股一手金额'
    sourcecol1 = [u'每手股数', u'招股价上限']
    computeexcelmoney(path, descol1, sourcecol1)
    computeexcelnum(path, descol, sourcecol)