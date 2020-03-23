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
        'Connection': 'keep-alive',
        'Cookie': 'UM_distinctid=16f03aeda951f7-0e445c640904db-12316b5a-13c680-16f03aeda96100; scopus.machineID=42A3D9C620FA39CD631B1EEC603CF6C6.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; __aza_perm=CheckPermissionCookie; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; xmlHttpRequest=true; everest_g_v2=g_surferid~XhQr8gAAFPW8EsFc; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; BIDUPSID=AD4E635AC5395DE343B6AC26FE46C873; PSTM=1580807141; BAIDUID=AD4E635AC5395DE32BF32FAB36D9518E:FG=1; _abck=DBA7C3BD927B70F835EE71C7C1AF1E5C~-1~YAAQ3tfSF1Al/c5vAQAALYYZEAP3PsVEykKgNX1Z7yztci5hDgw84xn4PLHgJApP9czi1665gl9UYf3Psj8hMewsC0VJyJ4zeYaAFnDRWBRn1V6QHFRN+AM1fdPLrdjVH03rIq8N/2uegBrLK1oqeHfP7xekO3huLcT5DQyMv7AsPC8sgBewMpaTC0r3cZ/sEsFtnKPxLxmLcpMThl0AM4o5v29/QwTJkACVqtbhSrQLE82F7JdGj0DDrhRbGu2lv036FerPg1M9d2o85I5lgfJvg6xNIT42/pvnc0U=~-1~-1~-1; sp=b75e6496-3e3c-4c58-acd1-2132ecf4d70c; _hjid=941e1bc6-293c-457e-bc91-4f5fc262dc09; _sp_id.9639=2314dd85-afcb-477b-9e82-64f1b98214d6.1580818111.2.1580906938.1580818818.be135e72-19df-41f8-acff-175cc2874545; s_vi=[CS]v1|2F1F44F18515E7A7-600006E4436E96C5[CE]; demdex=31931497131425969634038488604199850452; __cfduid=d09434e0603d6a09f49dbe43c5fe05fe11583840046; SCSessionID=4B6F4FFE39540F00BE032B486ED769B2.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=0eb3ee64-203f-45f0-8; check=true; javaScript=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1266252805%7CMCIDTS%7C18337%7CMCMID%7C10178513813538045854600690115189903694%7CMCAAMLH-1584944330%7C11%7CMCAAMB-1584944330%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1584346730s%7CNONE%7CMCAID%7CNONE%7CMCCIDH%7C1246701492%7CMCSYNCSOP%7C411-18276%7CvVersion%7C4.4.1; screenInfo="900:1440"; optimizelyDomainTestCookie=0.14344318356646446; optimizelyDomainTestCookie=0.13290482652625535; optimizelyDomainTestCookie=0.2357495105328169; mbox=PC#28144c226fe547c38b225004f77d29fb.22_0#1647584619|session#507372f3e07e413c9cd5562079e7073a#1584341384; s_pers=%20v8%3D1584339823873%7C1678947823873%3B%20v8_s%3DLess%2520than%25201%2520day%7C1584341623873%3B%20c19%3Dsc%253Asearch%253Adocument%2520results%7C1584341623931%3B%20v68%3D1584339817703%7C1584341624077%3B; s_sess=%20s_cpc%3D0%3B%20e78%3Dtitle%2528a%2520distributed%2520representation%2520approach%2520to%2520group%2520problem%2520solving%2529%3B%20c21%3Ddoi%252810.1002%252F%2528sici%25291097-4571%2528199807%252949%253A9%253C801%253A%253Aaid-asi5%253E3.0.co%253B2-q%2529%3B%20e13%3Ddoi%252810.1002%252F%2528sici%25291097-4571%2528199807%252949%253A9%253C801%253A%253Aaid-asi5%253E3.0.co%253B2-q%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20e41%3D1%3B%20s_sq%3D%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Asearch%25253Adocument%252520results%252C45%252C45%252C812%252C1440%252C246%252C1440%252C900%252C1%252CP%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520results%252C45%252C45%252C812%252C1440%252C246%252C1440%252C900%252C1%252CP%3B; _webvpn_key=eyJhbGciOiJIUzI1NiJ9.eyJ1c2VyIjoiMjAxNzIyNDQ5MiIsImlhdCI6MTU4NDM0Njg3NywiZXhwIjoxNTg0NDMzMjc3fQ.Tz7MjWdXGmj9RPXICff-MiRrXH1eOntkWdKJiUvX4j0; webvpn_username=2017224492%7C1584346877%7C0fe6f617fa5351d527cd5d4daca51d39d4471a77',
        'Host': 'www-scopus-com-s.webvpn.nefu.edu.cn',
        # 'Referer': 'https://www-scopus-com-s.webvpn.nefu.edu.cn/search/form.uri?display=basic&zone=header&origin=',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36',
}

SNIPheaders = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language':'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
        'Connection': 'keep-alive',
        'Cookie': 'UM_distinctid=16f03aeda951f7-0e445c640904db-12316b5a-13c680-16f03aeda96100; scopus.machineID=42A3D9C620FA39CD631B1EEC603CF6C6.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; __aza_perm=CheckPermissionCookie; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; xmlHttpRequest=true; everest_g_v2=g_surferid~XhQr8gAAFPW8EsFc; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; BIDUPSID=AD4E635AC5395DE343B6AC26FE46C873; PSTM=1580807141; BAIDUID=AD4E635AC5395DE32BF32FAB36D9518E:FG=1; _abck=DBA7C3BD927B70F835EE71C7C1AF1E5C~-1~YAAQ3tfSF1Al/c5vAQAALYYZEAP3PsVEykKgNX1Z7yztci5hDgw84xn4PLHgJApP9czi1665gl9UYf3Psj8hMewsC0VJyJ4zeYaAFnDRWBRn1V6QHFRN+AM1fdPLrdjVH03rIq8N/2uegBrLK1oqeHfP7xekO3huLcT5DQyMv7AsPC8sgBewMpaTC0r3cZ/sEsFtnKPxLxmLcpMThl0AM4o5v29/QwTJkACVqtbhSrQLE82F7JdGj0DDrhRbGu2lv036FerPg1M9d2o85I5lgfJvg6xNIT42/pvnc0U=~-1~-1~-1; sp=b75e6496-3e3c-4c58-acd1-2132ecf4d70c; _hjid=941e1bc6-293c-457e-bc91-4f5fc262dc09; _sp_id.9639=2314dd85-afcb-477b-9e82-64f1b98214d6.1580818111.2.1580906938.1580818818.be135e72-19df-41f8-acff-175cc2874545; s_vi=[CS]v1|2F1F44F18515E7A7-600006E4436E96C5[CE]; demdex=31931497131425969634038488604199850452; __cfduid=d09434e0603d6a09f49dbe43c5fe05fe11583840046; _webvpn_key=eyJhbGciOiJIUzI1NiJ9.eyJ1c2VyIjoiMjAxNzIyNDQ5MiIsImlhdCI6MTU4NDMzOTQ5MCwiZXhwIjoxNTg0NDI1ODkwfQ.U84NjgiSQtHvZOel_HCsvffFSeh-b2pXPY4g7sahnu4; webvpn_username=2017224492%7C1584339490%7C53db685781180a58f8aac8d034829dbe58d685b2; SCSessionID=4B6F4FFE39540F00BE032B486ED769B2.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=0eb3ee64-203f-45f0-8; check=true; javaScript=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1266252805%7CMCIDTS%7C18337%7CMCMID%7C10178513813538045854600690115189903694%7CMCAAMLH-1584944330%7C11%7CMCAAMB-1584944330%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1584346730s%7CNONE%7CMCAID%7CNONE%7CMCCIDH%7C1246701492%7CMCSYNCSOP%7C411-18276%7CvVersion%7C4.4.1; optimizelyDomainTestCookie=0.7528090685676205; optimizelyDomainTestCookie=0.9693966205348166; optimizelyDomainTestCookie=0.8988600813496952; mbox=PC#28144c226fe547c38b225004f77d29fb.22_0#1647584404|session#507372f3e07e413c9cd5562079e7073a#1584341384; screenInfo="900:1440"; optimizelyPendingLogEvents=%5B%5D; s_pers=%20c19%3Dsc%253Asearch%253Adocument%2520results%7C1584341410145%3B%20v68%3D1584339602533%7C1584341410279%3B%20v8%3D1584339701959%7C1678947701959%3B%20v8_s%3DLess%2520than%25201%2520day%7C1584341501959%3B; s_sess=%20c21%3Dtitle%2528a%2520distributed%2520representation%2520approach%2520to%2520group%2520problem%2520solving%2529%3B%20e13%3Dtitle%2528a%2520distributed%2520representation%2520approach%2520to%2520group%2520problem%2520solving%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20e41%3D1%3B%20s_cpc%3D0%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Asearch%25253Adocument%252520results%252C14%252C14%252C246%252C1440%252C246%252C1440%252C900%252C1%252CP%3B%20s_sq%3Delsevier-sc-prod%25252Celsevier-global-prod%253D%252526c.%252526a.%252526activitymap.%252526page%25253Dsc%2525253Asearch%2525253Adocument%25252520results%252526link%25253DJournal%25252520of%25252520the%25252520American%25252520Society%25252520for%25252520Information%25252520Science%252526region%25253DresultDataRow0%252526pageIDType%25253D1%252526.activitymap%252526.a%252526.c%252526pid%25253Dsc%2525253Asearch%2525253Adocument%25252520results%252526pidt%25253D1%252526oid%25253Dhttps%2525253A%2525252F%2525252Fwww-scopus-com-s.webvpn.nefu.edu.cn%2525252Fsourceid%2525252F36143%2525253Forigin%2525253Dresultslist%252526ot%25253DA%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520results%252C35%252C14%252C624%252C1440%252C246%252C1440%252C900%252C1%252CP%3B',
        'Host': 'www-scopus-com-s.webvpn.nefu.edu.cn',
        #'Referer': 'https://www-scopus-com-s.webvpn.nefu.edu.cn/results/results.uri?numberOfFields=0&src=s&clickedLink=&edit=&editSaveSearch=&origin=searchbasic&authorTab=&affiliationTab=&advancedTab=&scint=1&menu=search&tablin=&searchterm1=10.1007%2Fs11192-013-1004-x&field1=DOI&dateType=Publication_Date_Type&yearFrom=Before+1960&yearTo=Present&loadDate=7&documenttype=All&accessTypes=All&resetFormLink=&st1=10.1007%2Fs11192-013-1004-x&st2=&sot=b&sdt=b&sl=30&s=DOI%2810.1007%2Fs11192-013-1004-x%29&sid=870f69ab1c0c418086751b5d8cc1fa06&searchId=870f69ab1c0c418086751b5d8cc1fa06&txGid=bc614ae81986437d3280736de3f78ea0&sort=plf-f&originationType=b&rr=',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.100 Safari/537.36',

}

url1 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/results/results.uri?numberOfFields=0&src=s&clickedLink=&edit=&editSaveSearch=&origin=searchbasic&authorTab=&affiliationTab=&advancedTab=&scint=1&menu=search&tablin=&searchterm1='
url2 = '&field1=DOI&dateType=Publication_Date_Type&yearFrom=Before+1960&yearTo=Present&loadDate=7&documenttype=All&accessTypes=All&resetFormLink=&st1='
url3 = '&st2=&sot=b&sdt=b&sl=22&s=DOI%28'
url4 = '%29&sid=04d9016932b494f613c131f956db3e87&searchId=04d9016932b494f613c131f956db3e87&txGid=cd4965c8585c87e1b28188380fa2685e&sort=plf-f&originationType=b&rr='

citationturl1 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/search/submit/citedby.uri?eid='
citationturl2 = '&src=s&origin=recordpage'

snipurl = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/api/rest/sources/'

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

def get_main_words(s):
    s = serch(s, headers)
    a = s.findAll(name="div", attrs={"class": "sciTopicsVal displayNone"})
    # print(a)
    if len(a) != 0 and a[0].text != '':
        j = json.loads(str(a[0].text))
        return j['topic']['name']
    return ''

def get_H(s):
    s = s.find(name="span", attrs={"class": "ddmAuthorList"})
    hrefs = []
    H = []

    for i in s.children:
        if isinstance(i, bs4.element.Tag):
            hrefs.append(i.get('href'))

    for href in hrefs:
        s = serch(href, headers)
        s = s.find(name='section', attrs={"id": "authorDetailsHindex"}).find(name='span',attrs={"class": "fontLarge"})
        H.append(float(s.text))
    return H

def get_SNIP(s):
    r = requests.get(url=s, headers=SNIPheaders, allow_redirects=False)
    # print(r.text)
    data0 = re.findall(r'name>SNIP&lt;/name>(.*?)/value>', r.text, re.S)
    data = re.findall(r'&lt;value>(.*?)&lt;', str(data0), re.S)
    # data = re.findall(r'&lt;value>(.*?)&lt;/value', r.text, re.S)
    if data:
        snip = data[0]
        return snip
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
    for i in range(437,438):
        paper = {}
        content = table.row_values(i)
        paper['标题'] = content[2]
        paper['年份'] = content[3]
        paper['来源出版物名称'] = content[4]
        paper['DOI'] = content[12]
        paper['EID'] = content[42]
        print(paper['标题'])
        print(paper['年份'])
        print(paper['来源出版物名称'])
        print(paper['DOI'])
        print(paper['EID'])
        papers.append(paper)
    return papers

if __name__ == '__main__':

    papers = get_excel()
    fSub = open('./subject.txt', 'w')
    fCoun = open('./country.txt', 'w')
    # fHIdex = open('./hIdex.txt', 'w')
    # fCitation = open('./citation.txt', 'w')
    # fMainWords = open('./mainWords.txt', 'w')
    # fSnip = open('./snip.txt', 'w')
    fAffi = open('./affiliation.txt', 'w')


    for paper in papers:
        eid = paper['EID']
        doi = paper['DOI']

        # ====================================1=====================================

        if doi != '':
            url = url1 + doi + url2 + doi + url3 + doi + url4
            page_source = requests.get(url=url, headers=headers, allow_redirects=False)
            print(page_source.text)

            # -----------------学科-------------------

            subjectArea = get_subjectArea(page_source.text)
            print("Subject area:")
            print(subjectArea)
            print("\n")

            print(subjectArea, file=fSub, flush=True)
            # -----------------国家-------------------

            print("Country:")
            country = get_country(page_source.text)
            print(country)
            print("\n")
            print(country, file=fCoun, flush=True)
            time.sleep(1)
        else:
            print('', file=fSub, flush=True)
            print('', file=fCoun, flush=True)


        # ====================================2=====================================

        # if doi != '':
        #
        #     url = url1 + doi + url2 + doi + url3 + doi + url4
        #     s = serch(url, headers)
        #
        #     # # -----------------H指数-------------------
        #     # H = get_H(s)
        #     # print(H)
        #     # if H != []:
        #     #     maxH = max(H)
        #     #     print(int(maxH), file=fHIdex, flush=True)
        #     #     print("H Index:")
        #     #     print(maxH)
        #     #     print("\n")
        #     # else:
        #     #     print('', file=fHIdex, flush=True)
        #     # time.sleep(1)
        #
        #
        #     # -----------------主题词-------------------
        #     try:
        #         mainurl = s.find('a', title="Show document details").get('href')
        #     except:
        #         print('', file=fMainWords, flush=True)
        #     else:
        #         mainWords = get_main_words(mainurl)
        #         print(mainWords, file=fMainWords, flush=True)
        #         print("Main Words:")
        #         print(mainWords)
        #         print("\n")
        #     time.sleep(1)
        #
        #     # -------------------snip-------------------
        #     try:
        #         snipurl0 = s.find('a', title="Show source title details").get('href')
        #     except:
        #         print('', file=fSnip, flush=True)
        #     else:
        #         snipurl0 = str(snipurl0).split('?')[0].split('/')[-1]
        #         snipurl0 = snipurl + snipurl0
        #         snip = get_SNIP(snipurl0)
        #         print(snip, file=fSnip, flush=True)
        #         print("Snip:")
        #         print(snip)
        #         print("\n")
        #     time.sleep(1)
        # else:
        #     # print('', file=fHIdex, flush=True)
        #     print('', file=fMainWords, flush=True)
        #     print('', file=fSnip, flush=True)

        # ====================================3=====================================

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

        # ====================================4=====================================

        if doi != '':
            # -----------------机构--------------------

            s = 'DOI(' + doi + ')'
            st1 = doi
            data = {
                'clusterDisplayCount': '10',
                'sot': 'b',
                'navigatorName': 'AFFIL',
                'clusterCategory': 'selectedAffiliationClusterCategories',
                'cite': '',
                'refeid': '',
                'refeidnss': '',
                's': s,
                'st1': st1,
                'st2': '',
                'sid': 'e635e35a50254e190a9379ccc39a7b30',
                'sdt': 'b',
                'sort': 'plf-f',
                'citingId': '',
                'citedAuthorId': '',
                'listId': '',
                'origin': 'resultslist',
                'src': 's',
                'affilCity': '',
                'affilName': '',
                'affilCntry': '',
                'affiliationId': '',
                'cluster': '',
                'offset': '1',
                'scla': '',
                'scls': '',
                'sclk': '',
                'scll': '',
                'sclsb': '',
                'sclc': '',
                'scfs': '',
                'ref': '',
                'isRebrandLayout': 'true',
            }
            rep = requests.post(
                url='https://www-scopus-com-s.webvpn.nefu.edu.cn/standard/retrieveClusterAttributes.uri', data=data,
                headers=headers)
            affiliation = re.findall(r'class="btnText">(.*?)</span>', rep.text, re.S)
            print("\nAffiliation:")
            print(affiliation)
            print(affiliation, file=fAffi, flush=True)
            time.sleep(1)
        else:
            print('', file=fAffi, flush=True)



        print("=======================================================================")
