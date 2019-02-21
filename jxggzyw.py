import requests
from bs4 import BeautifulSoup
import lxml
from lxml import etree
import xlwt
import json
import jsonpath
import os


header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.106 Safari/537.36','referer':'http://www.jxsggzy.cn/web/tradeSubject_list.html','Cookie':'JSESSIONID=6E8A83AF7E01C129FF5B91317771B3AC; _CSRFCOOKIE=C040DA20A104FA8DA84ECB1C1B687FAFEDABC045','Host':'www.jxsggzy.cn'}
url = 'http://www.jxsggzy.cn/jxggzy/services/JyxxWebservice/getTradeList?response=application/json&pageIndex=2&pageSize=22&&dsname=ztb_data&bname=&qytype=3&itemvalue=131'

def getHtmlText(url,header):
    try:
        #url = 'http://bykiss.club/tophot.html'
        html = requests.get(url,headers=header)
        html.raise_for_status()
        html.encoding = html.apparent_encoding
        #print(html.raise_for_status())
        #print(html.text)
        return html.text
    except:
        print('访问失败')
def getListinfo(html):
    jsonData = json.loads(html)
    getData = jsonData.get('return')
    data = json.loads(getData)
    Listinfo = []
    i = 0
    getTable = data.get('Table')
    for List in getTable:
        alinkList = getTable[i].get('alink')
        szdqList = getTable[i].get('szdq')
        qymcList = getTable[i].get('qymc')
        Listinfo.append(alinkList)
        Listinfo.append(szdqList)
        Listinfo.append(qymcList)
        i += 1
    return Listinfo
def saveToExcel(Listinfo,countPage):
    workbook=xlwt.Workbook()
    sheet1=workbook.add_sheet('sheet1',cell_overwrite_ok=True)
    k=0
    for i in range(countPage):
        for j in range(3):
            print('正在写入的行和列是',i,j)
            sheet1.write(i,j,Listinfo[k])
            k+=1
    name = str(input('请输入保存文件名：'))
    path = os.getcwd() + '\\'
    workbook.save(path + name + '.xls')
    print('你保存文件的位置：',path + name +'.xls')
    print('任务已结束')
def getUrl(page):
    url = ''
    onepage = str(page)
    headUrl = 'http://www.jxsggzy.cn/jxggzy/services/JyxxWebservice/getTradeList?response=application/json&pageIndex='
    footUrl = '&pageSize=22&&dsname=ztb_data&bname=&qytype=3&itemvalue=131'
    url = headUrl + onepage + footUrl
    return url


def main():
    header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.106 Safari/537.36',
        'referer': 'http://www.jxsggzy.cn/web/tradeSubject_list.html',
        'Cookie': 'JSESSIONID=6E8A83AF7E01C129FF5B91317771B3AC; _CSRFCOOKIE=C040DA20A104FA8DA84ECB1C1B687FAFEDABC045',
        'Host': 'www.jxsggzy.cn'}
    allList = []
    starPage = input(" 请输入开始页码(1-323)：")
    endPage = int(input(" 请输入结束页码(1-323)："))
    page = int(starPage)
    countPage = (endPage-int(starPage)+1)*22
    while (page<=endPage):
        allList+=getListinfo(getHtmlText(getUrl(page), header))
        page+=1
    print('共计',countPage,'条')
    saveToExcel(allList,countPage)
    #print(allList)

if __name__=='__main__':
    main()



