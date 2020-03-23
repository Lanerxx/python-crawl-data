import urllib.request
from bs4 import BeautifulSoup
import bs4
import requests
import json
import io
import locale
import xlrd
import time
import re
import sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='gb18030')

headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Cookie':'UM_distinctid=16f03aeda951f7-0e445c640904db-12316b5a-13c680-16f03aeda96100; scopus.machineID=42A3D9C620FA39CD631B1EEC603CF6C6.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; __aza_perm=CheckPermissionCookie; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; xmlHttpRequest=true; everest_g_v2=g_surferid~XhQr8gAAFPW8EsFc; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; BIDUPSID=AD4E635AC5395DE343B6AC26FE46C873; PSTM=1580807141; BAIDUID=AD4E635AC5395DE32BF32FAB36D9518E:FG=1; _abck=DBA7C3BD927B70F835EE71C7C1AF1E5C~-1~YAAQ3tfSF1Al/c5vAQAALYYZEAP3PsVEykKgNX1Z7yztci5hDgw84xn4PLHgJApP9czi1665gl9UYf3Psj8hMewsC0VJyJ4zeYaAFnDRWBRn1V6QHFRN+AM1fdPLrdjVH03rIq8N/2uegBrLK1oqeHfP7xekO3huLcT5DQyMv7AsPC8sgBewMpaTC0r3cZ/sEsFtnKPxLxmLcpMThl0AM4o5v29/QwTJkACVqtbhSrQLE82F7JdGj0DDrhRbGu2lv036FerPg1M9d2o85I5lgfJvg6xNIT42/pvnc0U=~-1~-1~-1; sp=b75e6496-3e3c-4c58-acd1-2132ecf4d70c; _hjid=941e1bc6-293c-457e-bc91-4f5fc262dc09; _sp_id.9639=2314dd85-afcb-477b-9e82-64f1b98214d6.1580818111.2.1580906938.1580818818.be135e72-19df-41f8-acff-175cc2874545; s_vi=[CS]v1|2F1F44F18515E7A7-600006E4436E96C5[CE]; demdex=31931497131425969634038488604199850452; __cfduid=d09434e0603d6a09f49dbe43c5fe05fe11583840046; _webvpn_key=eyJhbGciOiJIUzI1NiJ9.eyJ1c2VyIjoiMjAxNzIyNDQ5MiIsImlhdCI6MTU4NDk2NzU3MCwiZXhwIjoxNTg1MDUzOTcwfQ.E7-8cv7KjwRoQ3YTPw5FMSI_Gj5VJZU6915_GsDItso; webvpn_username=2017224492%7C1584967570%7C50addbaf3f9ca70f96274fe9dbc1100f35f1fa27; check=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1266252805%7CMCIDTS%7C18345%7CMCMID%7C10178513813538045854600690115189903694%7CMCAAMLH-1585572402%7C11%7CMCAAMB-1585572402%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1584974802s%7CNONE%7CMCAID%7CNONE%7CMCCIDH%7C1246701492%7CMCSYNCSOP%7C411-18276%7CvVersion%7C4.4.1; javaScript=true; screenInfo="900:1440"; SCSessionID=9C06A84AABA0FAA003067A688EAE11EC.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=87082e4d-aa89-4a39-9; optimizelyDomainTestCookie=0.0843606951783824; optimizelyDomainTestCookie=0.4418508759822024; optimizelyDomainTestCookie=0.10494631451816372; mbox=PC#28144c226fe547c38b225004f77d29fb.22_0#1648216330|session#0e5c9fd0fdf24aa381c8b4e65bf21201#1584973295; s_pers=%20v8%3D1584971538385%7C1679579538385%3B%20v8_s%3DMore%2520than%25207%2520days%7C1584973338385%3B%20c19%3Dsc%253Asearch%253Adocument%2520searchform%7C1584973338440%3B%20v68%3D1584971532556%7C1584973338487%3B; s_sess=%20s_cpc%3D0%3B%20c21%3Dtitle-abs-key%2528color%2529%3B%20e13%3Dtitle-abs-key%2528color%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20e78%3Dtitle-abs-key%2528color%2529%3B%20s_sq%3D%3B%20s_cc%3Dtrue%3B%20e41%3D1%3B%20s_ppvl%3Dsc%25253Arecord%25253Asource%252520info%252C20%252C20%252C311%252C1440%252C167%252C1440%252C900%252C1%252CP%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520searchform%252C37%252C37%252C387%252C1440%252C202%252C1440%252C900%252C1%252CP%3B',
        'Host': 'www-scopus-com-s.webvpn.nefu.edu.cn',
        # 'Referer': 'https://www-scopus-com-s.webvpn.nefu.edu.cn/search/form.uri?display=basic&zone=header&origin=',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36',
}

headerDetail = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Cookie':'scopus.machineID=2836B351119B39A9DCCB7F368064A138.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1584860302365r0.0014021053918049642; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; optimizelyEndUserId=oeu1584860302365r0.0014021053918049642; javaScript=true; uuid2=1407715896380496685; check=true; demdex=91224598901239483393216475646940455076; xmlHttpRequest=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; everest_g_v2=g_surferid~XncM7wAAAsFAYJpK; dpm=91224598901239483393216475646940455076; screenInfo="1440:2560"; JSESSIONID=e6a48116cba9e768; __cfruid=a314e85176bd63143354fa56e43280236f6fe028-1584860453; __aza_perm=CheckPermissionCookie; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1266252805%7CMCIDTS%7C18344%7CMCMID%7C91214671505630802303220000527927876017%7CMCAAMLH-1585472424%7C11%7CMCAAMB-1585472424%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1584874824s%7CNONE%7CMCAID%7CNONE%7CMCCIDH%7C1246701492%7CMCSYNCSOP%7C411-18351%7CvVersion%7C4.4.1; __cfduid=da610730c4c275504994b583a57cfdd8d1584868370; SCSessionID=3DE406137CA01B0367CA89143AF4A67E.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=a288f842-2768-40c3-8; _webvpn_key=eyJhbGciOiJIUzI1NiJ9.eyJ1c2VyIjoiMjAxNzIyNDQ5MiIsImlhdCI6MTU4NDg3MDAwNSwiZXhwIjoxNTg0OTU2NDA1fQ.klZYU9eelGMmyQ_b_eD3UcdFM2uBxDHqxUDCg7gcsCU; webvpn_username=2017224492%7C1584870005%7C1957461299ccb36d81a5cb1542b124cfaa6129b4; anj=dTM7k!M4/8F7/.XF\']wIg2GU#izRZL!_13mM-SAOo>r`fD$25ADb6`NDYtLArFM*I; mbox=PC#529e791d93d0450faa75ea90ee68505c.22_0#1648114910|session#81b77916d40d4b88984ef9a41b98f623#1584871893; s_pers=%20v8%3D1584870111233%7C1679478111233%3B%20v8_s%3DFirst%2520Visit%7C1584871911233%3B%20c19%3Dsc%253Arecord%253Aauthor%2520details%7C1584871911242%3B%20v68%3D1584870107434%7C1584871911257%3B; s_sess=%20s_cpc%3D0%3B%20c21%3Dtitle%2528effect%2520of%2520cooperation%2520between%2520chinese%2520scientific%2520journals%2520and%2520international%2520publishers%2520on%2520journals%2520impact%2520factor%2529%3B%20e13%3Dtitle%2528effect%2520of%2520cooperation%2520between%2520chinese%2520scientific%2520journals%2520and%2520international%2520publishers%2520on%2520journals%2520impact%2520factor%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20e78%3Dtitle%2528effect%2520of%2520cooperation%2520between%2520chinese%2520scientific%2520journals%2520and%2520international%2520publishers%2520on%2520journals%2527%2520impact%2520factor%2529%3B%20s_sq%3D%3B%20e41%3D1%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Arecord%25253Aauthor%252520details%252C44%252C44%252C1273%252C1189%252C1273%252C2560%252C1440%252C1%252CL%3B%20s_ppv%3Dsc%25253Arecord%25253Adocument%252520record%252C18%252C18%252C818%252C1189%252C818%252C2560%252C1440%252C1%252CP%3B',
        'Host': 'www-scopus-com-s.webvpn.nefu.edu.cn',
        # 'Referer': 'https://www-scopus-com-s.webvpn.nefu.edu.cn/search/form.uri?display=basic&zone=header&origin=',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36',
}

SNIPheaders = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
        'Connection': 'keep-alive',
        'Cookie': 'UM_distinctid=16f03aeda951f7-0e445c640904db-12316b5a-13c680-16f03aeda96100; scopus.machineID=42A3D9C620FA39CD631B1EEC603CF6C6.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; __aza_perm=CheckPermissionCookie; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; xmlHttpRequest=true; everest_g_v2=g_surferid~XhQr8gAAFPW8EsFc; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; BIDUPSID=AD4E635AC5395DE343B6AC26FE46C873; PSTM=1580807141; BAIDUID=AD4E635AC5395DE32BF32FAB36D9518E:FG=1; _abck=DBA7C3BD927B70F835EE71C7C1AF1E5C~-1~YAAQ3tfSF1Al/c5vAQAALYYZEAP3PsVEykKgNX1Z7yztci5hDgw84xn4PLHgJApP9czi1665gl9UYf3Psj8hMewsC0VJyJ4zeYaAFnDRWBRn1V6QHFRN+AM1fdPLrdjVH03rIq8N/2uegBrLK1oqeHfP7xekO3huLcT5DQyMv7AsPC8sgBewMpaTC0r3cZ/sEsFtnKPxLxmLcpMThl0AM4o5v29/QwTJkACVqtbhSrQLE82F7JdGj0DDrhRbGu2lv036FerPg1M9d2o85I5lgfJvg6xNIT42/pvnc0U=~-1~-1~-1; sp=b75e6496-3e3c-4c58-acd1-2132ecf4d70c; _hjid=941e1bc6-293c-457e-bc91-4f5fc262dc09; _sp_id.9639=2314dd85-afcb-477b-9e82-64f1b98214d6.1580818111.2.1580906938.1580818818.be135e72-19df-41f8-acff-175cc2874545; s_vi=[CS]v1|2F1F44F18515E7A7-600006E4436E96C5[CE]; demdex=31931497131425969634038488604199850452; __cfduid=d09434e0603d6a09f49dbe43c5fe05fe11583840046; _webvpn_key=eyJhbGciOiJIUzI1NiJ9.eyJ1c2VyIjoiMjAxNzIyNDQ5MiIsImlhdCI6MTU4NDk2NzU3MCwiZXhwIjoxNTg1MDUzOTcwfQ.E7-8cv7KjwRoQ3YTPw5FMSI_Gj5VJZU6915_GsDItso; webvpn_username=2017224492%7C1584967570%7C50addbaf3f9ca70f96274fe9dbc1100f35f1fa27; check=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1266252805%7CMCIDTS%7C18345%7CMCMID%7C10178513813538045854600690115189903694%7CMCAAMLH-1585572402%7C11%7CMCAAMB-1585572402%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1584974802s%7CNONE%7CMCAID%7CNONE%7CMCCIDH%7C1246701492%7CMCSYNCSOP%7C411-18276%7CvVersion%7C4.4.1; javaScript=true; screenInfo="900:1440"; SCSessionID=D32A0416AF326B80824BCC3AD1C7C5A1.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=a5b1c880-9559-4448-9; s_pers=%20v8%3D1584969946598%7C1679577946598%3B%20v8_s%3DMore%2520than%25207%2520days%7C1584971746598%3B%20c19%3Dsc%253Asearch%253Adocument%2520results%7C1584971746650%3B%20v68%3D1584969927280%7C1584971746861%3B; s_sess=%20s_cpc%3D0%3B%20e78%3Ddoi%252810.1016%252Fj.joi.2010.05.002%2529%3B%20s_sq%3D%3B%20c21%3Ddoi%252810.1016%252Fj.joi.2010.05.002%2529%3B%20e13%3Ddoi%252810.1016%252Fj.joi.2010.05.002%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20s_ppvl%3Dsc%25253Asearch%25253Adocument%252520results%252C30%252C30%252C566%252C1440%252C152%252C1440%252C900%252C1%252CP%3B%20e41%3D1%3B%20s_cc%3Dtrue%3B%20s_ppv%3Dsc%25253Arecord%25253Asource%252520info%252C51%252C10%252C789%252C1440%252C173%252C1440%252C900%252C1%252CP%3B; optimizelyPendingLogEvents=%5B%5D; optimizelyDomainTestCookie=0.5256868940307318; optimizelyDomainTestCookie=0.8893264344213416; optimizelyDomainTestCookie=0.5467865386809638; mbox=PC#28144c226fe547c38b225004f77d29fb.22_0#1648214760|session#a0e138a2778249d3b607a3b7461a9031#1584971327',
        'Host': 'www-scopus-com-s.webvpn.nefu.edu.cn',
        # 'Referer': https://www-scopus-com-s.webvpn.nefu.edu.cn/sourceid/5100155103?origin=resultslist
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Moz\'illa/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36',
        'X-NewRelic-ID': 'VQQPUFdVCRADVVVXAwABVA==',
        'X-Requested-With': 'XMLHttpRequest',
}

url1 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/results/results.uri?numberOfFields=0&src=s&clickedLink=&edit=&editSaveSearch=&origin=searchbasic&authorTab=&affiliationTab=&advancedTab=&scint=1&menu=search&tablin=&searchterm1='
url2 = '&field1=DOI&dateType=Publication_Date_Type&yearFrom=Before+1960&yearTo=Present&loadDate=7&documenttype=All&accessTypes=All&resetFormLink=&st1='
url3 = '&st2=&sot=b&sdt=b&sl=22&s=DOI%28'
url4 = '%29&sid=04d9016932b494f613c131f956db3e87&searchId=04d9016932b494f613c131f956db3e87&txGid=cd4965c8585c87e1b28188380fa2685e&sort=plf-f&originationType=b&rr='

citationturl1 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/search/submit/citedby.uri?eid='
citationturl2 = '&src=s&origin=recordpage'

snipurl = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/api/rest/sources/'

detail1 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/record/display.uri?eid='
# 2-s2.0-84944149175
detail2 = '&origin=resultslist&sort=plf-f&src=s&st1='
# 10.1007
detail3 = '%2f'
# s11192-014-1269-8
detail4 = '&st2=&sid=1e6d0b419dd875287a224a2fbc71ea74&sot=b&sdt=b&sl=30&s=DOI%2810.'
# 1007
detail5 = '%2f'
# s11192-014-1269-8
detail6 = '%29&relpos=0&citeCnt=46&searchTerm='

snip1 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/results/results.uri?numberOfFields=0&src=s&clickedLink=&edit=&editSaveSearch=&origin=searchbasic&authorTab=&affiliationTab=&advancedTab=&scint=1&menu=search&tablin=&searchterm1='
# 10.1016%2Fj.joi.2009.11.002
snip2 = '&field1=DOI&dateType=Publication_Date_Type&yearFrom=Before+1960&yearTo=Present&loadDate=7&documenttype=All&accessTypes=All&resetFormLink=&st1='
# 10.1016%2Fj.joi.2009.11.002
snip3 = '&st2=&sot=b&sdt=b&sl=30&s=DOI%28'
# 10.1016%2Fj.joi.2009.11.002
snip4 = '%29&sid=f90ade8a461ec3d0e4486e0fb8eb8e48&searchId=f90ade8a461ec3d0e4486e0fb8eb8e48&txGid=7d30969d45fa0773883a37d730690f93&sort=plf-f&originationType=b&rr='
# snip web
snip5 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/api/rest/sources/'


def serch(url,headers):
    req = urllib.request.Request(url=url, headers=headers)
    rsp = urllib.request.urlopen(req,timeout=20000)
    html = rsp.read().decode()
    s = BeautifulSoup(html, 'html.parser')
    return s

def get_2016_2017(s):
    yin2016 = 0
    yin2017 = 0
    have2016 = s.find(name="li", attrs={"id": "li_2016"})

    if have2016 != None:
        yin2016 = have2016.find(name='span', attrs={'class': 'badge'}).find(name='span', attrs={'class': 'btnText'})
        # print(yin2016)
        yin2016 = locale.atoi(yin2016.text.replace(',',''))

    have2017 = s.find(name="li", attrs={"id": "li_2017"})
    if have2017 != None:
        yin2017 = have2017.find(name='span', attrs={'class': 'badge'}).find(name='span', attrs={'class': 'btnText'})
        yin2017 = locale.atoi(yin2017.text.replace(',',''))
    return yin2016+yin2017

def get_main_words(url):
    s = requests.Session()
    page_source = s.get(url, headers=headerDetail)
    mainWords0= re.findall(r'<div class="sciTopicsVal displayNone"(.*?)</div>', page_source.text, re.S)
    mainWords = re.findall(r'"name":"(.*?)","id', str(mainWords0), re.S)
    return mainWords

def get_H(url):
    s = requests.Session()
    page_source = s.get(url, headers=headerDetail)
    H = []
    hIndexUrls0 = re.findall(r'<section id="authorlist">(.*?)</section>', page_source.text, re.S)
    hIndexUrls = re.findall(r'"name":(.*?)" title="显示作者详情"', str(hIndexUrls0), re.S)
    print(hIndexUrls)
    for hIndexUrl in hIndexUrls:
        print(hIndexUrl)
        s = requests.Session()
        hIndexUrl_source = s.get(hIndexUrl, headers=headerDetail)
        hIndex0 = re.findall(r'<span class="spanItalic">h</span>-Index:(.*?)<button type="button', hIndexUrl_source.text, re.S)
        hIndex = re.findall(r'<div><span class="fontLarge">(.*?)</span>', str(hIndex0), re.S)
        if hIndex:
            H.append(hIndex[0])
    return H

def get_SnipSjrRp(url):
    dataSnipSjrRp = []
    s = requests.Session()
    page_source = s.get(url, headers=headers)
    data0 = re.findall(r'<td data-type="source">\n<a href="(.*?)"  title="Show source title details"  class="ddmDocSource"  id="sourceTitle" >', page_source.text, re.S)
    print(data0)
    if data0:
        data1 = data0[0]
        data2 = data1[10:20]
        print(data2)
        snipUrl = snip5 + data2
        print(snipUrl)
        s1 = requests.Session()
        page_source1 = s1.get(snipUrl, headers=headers)
        datasnip0 = re.findall(r'name>SNIP&lt;/name>(.*?)/value>', page_source1.text, re.S)
        datasnip = re.findall(r'&lt;value>(.*?)&lt;', str(datasnip0), re.S)
        if datasnip:
            dataSnipSjrRp.append(datasnip[0])
            print(dataSnipSjrRp)
        else:
            dataSnipSjrRp.append('')

        datasjr0 = re.findall(r'name>SJR&lt;/name>(.*?)/value>', page_source1.text, re.S)
        datasjr = re.findall(r'&lt;value>(.*?)&lt;', str(datasjr0), re.S)
        if datasjr:
            dataSnipSjrRp.append(datasjr[0])
            print(dataSnipSjrRp)
        else:
            dataSnipSjrRp.append('')

        datarp0 = re.findall(r'name>RP&lt;/name>(.*?)/value>', page_source1.text, re.S)
        datarp = re.findall(r'&lt;value>(.*?)&lt;', str(datarp0), re.S)
        if datarp:
            dataSnipSjrRp.append(datarp[0])
            print(dataSnipSjrRp)
        else:
            dataSnipSjrRp.append('')

        return dataSnipSjrRp

    else:
        return []


def get_subjectArea(page_source):
    subData0 = re.findall(r'<label class="checkbox-label" for=\'cat_SUBJAREA(.*?)\n</label>', page_source, re.S)
    subData = re.findall(r'class="btnText">\\n(.*?)\\n</span>', str(subData0), re.S)
    return subData

def get_country(page_source):
    counData0 = re.findall(r'<label class="checkbox-label" for=\'cat_COUNTRY(.*?)\n</label>', page_source, re.S)
    counData = re.findall(r'class="btnText">\\n(.*?)\\n</span>', str(counData0), re.S)
    return counData

def get_excel():
    file = "./1-5.xls"

    data = xlrd.open_workbook(file, formatting_info=True)
    # new_workbook = copy(data)
    # new_worksheet = new_workbook.get_sheet(0)
    # print(data)
    table = data.sheet_by_name('test1')
    papers = []
    for i in range(453,470):
        paper = {}
        content = table.row_values(i)
        paper['标题'] = content[2]
        paper['年份'] = content[3]
        paper['来源出版物名称'] = content[4]
        paper['DOI'] = content[12]
        paper['EID'] = content[42]
        print(paper['标题'])
        # print(paper['年份'])
        # print(paper['来源出版物名称'])
        # print(paper['DOI'])
        # print(paper['EID'])
        papers.append(paper)
    return papers

if __name__ == '__main__':

    papers = get_excel()
    # fSub = open('./subject.txt', 'w')
    # fCoun = open('./country.txt', 'w')
    # fHIdex = open('./hIdex.txt', 'w')
    # fCitation = open('./citation.txt', 'w')
    # fMainWords = open('./mainWords.txt', 'w')
    fSnip = open('./snip.txt', 'w')
    fSjr = open('./sjr.txt', 'w')
    fRp = open('./rp.txt', 'w')
    # fAffi = open('./affiliation.txt', 'w')

    snipSjrRpFlag = 0

    for paper in papers:
        eid = paper['EID']
        doi = paper['DOI']

        # ====================================1=====================================

        #
        # if doi != '':
        # url = url1 + doi + url2 + doi + url3 + doi + url4
        #     print(url)
        #     s = requests.Session()
        #     page_source = s.get(url, headers=headers)
        #     # page_source = requests.get(url=url, headers=headers, allow_redirects=False)
        #
        #     # -----------------学科-------------------
        #
        #     subjectArea = get_subjectArea(page_source.text)
        #     print("Subject area:")
        #     print(subjectArea)
        #     print("\n")
        #
        #     print(subjectArea, file=fSub, flush=True)
        #     # -----------------国家-------------------
        #
        #     print("Country:")
        #     country = get_country(page_source.text)
        #     print(country)
        #     print("\n")
        #     print(country, file=fCoun, flush=True)
        #     time.sleep(1)
        # else:
        #     print('', file=fSub, flush=True)
        #     print('', file=fCoun, flush=True)

        # ====================================2=====================================

        if doi != '' and eid !='':
            # doi1 = doi[0:7]
            # doi2 = doi[8:]
            # doi3 = doi[3:7]
            # url = detail1 + eid + detail2 + doi1 + detail3 + doi2 + detail4 + doi3  + detail5 + doi2  +detail6

            # # -----------------H指数-------------------
            # H = get_H(url)
            # print(H)
            # if H != []:
            #     maxH = max(H)
            #     print(int(maxH), file=fHIdex, flush=True)
            #     print("H Index:")
            #     print(maxH)
            #     print("\n")
            # else:
            #     print('', file=fHIdex, flush=True)
            # time.sleep(1)


            # # -----------------主题词-------------------

            # mainWords = get_main_words(url)
            # print(mainWords, file=fMainWords, flush=True)
            # print("Main Words:")
            # print(mainWords)
            # print("\n")
            # time.sleep(1)
            #

            if snipSjrRpFlag == 0:
                snipSjrRpFlag = snipSjrRpFlag + 1
                # # --------------SNIP SJR RP----------------
                # 利用snipSjrRpFlag标识，4收1策略
                doiSnip = doi.replace('/','%2F')
                urlSnip0 = snip1 + doiSnip + snip2 + doiSnip  + snip3  + doiSnip + snip4
                snipSjrRp = get_SnipSjrRp(urlSnip0)
                if snipSjrRp:
                    snip = snipSjrRp[0]
                    sjr = snipSjrRp[1]
                    rp = snipSjrRp[2]
                    print(snip, file=fSnip, flush=True)
                    print(sjr, file=fSjr, flush=True)
                    print(rp, file=fRp, flush=True)

                    print("SNIP:" + snip + "  SJR" + sjr +"  RJ" + rp)
                    print("\n")
                time.sleep(1)
            else:
                snipSjrRpFlag = snipSjrRpFlag + 1
                if snipSjrRpFlag == 3:
                    snipSjrRpFlag = 0
                print('', file=fSnip, flush=True)
                print('', file=fSjr, flush=True)
                print('', file=fRp, flush=True)

        else:
            print(1)
            # print('', file=fHIdex, flush=True)
            # print('', file=fMainWords, flush=True)
            print('', file=fSnip, flush=True)
            print('', file=fSjr, flush=True)
            print('', file=fRp, flush=True)



        # # ====================================3=====================================
        #
        # if eid != '':
        #     url = citationturl1 + eid + citationturl2
        #     s = serch(url, headers)
        #
        #     # --------------2016_2017引用次数---------------
        #     citation = get_2016_2017(s)
        #     print(citation, file=fCitation, flush=True)
        #     print("2016_2017:")
        #     print(citation)
        # else:
        #     print('', file=fCitation, flush=True)
        #
        # # ====================================4=====================================
        #
        # if doi != '':
        #     # -----------------机构--------------------
        #
        #     s = 'DOI(' + doi + ')'
        #     st1 = doi
        #     data = {
        #         'clusterDisplayCount': '10',
        #         'sot': 'b',
        #         'navigatorName': 'AFFIL',
        #         'clusterCategory': 'selectedAffiliationClusterCategories',
        #         'cite': '',
        #         'refeid': '',
        #         'refeidnss': '',
        #         's': s,
        #         'st1': st1,
        #         'st2': '',
        #         'sid': 'e635e35a50254e190a9379ccc39a7b30',
        #         'sdt': 'b',
        #         'sort': 'plf-f',
        #         'citingId': '',
        #         'citedAuthorId': '',
        #         'listId': '',
        #         'origin': 'resultslist',
        #         'src': 's',
        #         'affilCity': '',
        #         'affilName': '',
        #         'affilCntry': '',
        #         'affiliationId': '',
        #         'cluster': '',
        #         'offset': '1',
        #         'scla': '',
        #         'scls': '',
        #         'sclk': '',
        #         'scll': '',
        #         'sclsb': '',
        #         'sclc': '',
        #         'scfs': '',
        #         'ref': '',
        #         'isRebrandLayout': 'true',
        #     }
        #     rep = requests.post(
        #         url='https://www-scopus-com-s.webvpn.nefu.edu.cn/standard/retrieveClusterAttributes.uri', data=data,
        #         headers=headers)
        #     affiliation = re.findall(r'class="btnText">(.*?)</span>', rep.text, re.S)
        #     print("\nAffiliation:")
        #     print(affiliation)
        #     print(affiliation, file=fAffi, flush=True)
        #     time.sleep(1)
        # else:
        #     print('', file=fAffi, flush=True)
        #
        #

        print("=======================================================================")
