# ----------------------正经的-----------------------
import xlrd
from selenium import webdriver
import time
import requests
import re

url1 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/results/results.uri?numberOfFields=0&src=s&clickedLink=&edit=&editSaveSearch=&origin=searchbasic&authorTab=&affiliationTab=&advancedTab=&scint=1&menu=search&tablin=&searchterm1='
url2 = '&field1=DOI&dateType=Publication_Date_Type&yearFrom=Before+1960&yearTo=Present&loadDate=7&documenttype=All&accessTypes=All&resetFormLink=&st1='
url3 = '&st2=&sot=b&sdt=b&sl=22&s=DOI%28'
url4 = '%29&sid=04d9016932b494f613c131f956db3e87&searchId=04d9016932b494f613c131f956db3e87&txGid=cd4965c8585c87e1b28188380fa2685e&sort=plf-f&originationType=b&rr='

souUrl1 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/results/results.uri?sort=cp-t&src=s&nlo=&nlr=&nls=&sid=98e16c8a1c4c26fa3ab53304e1f23de8&sot=a&sdt=cl&cluster=scopubyr%2C%22'
# 2017
souUrl2 = '%22%2Ct&sl=17&s=SOURCE-ID+%28'
# 24222
souUrl3 = '%29&origin=resultslist&zone=leftSideBar&editSaveSearch=&txGid=3140c9fb2995db02cf46e4dba6b4c152'

headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Cookie':'UM_distinctid=16f03aeda951f7-0e445c640904db-12316b5a-13c680-16f03aeda96100; scopus.machineID=42A3D9C620FA39CD631B1EEC603CF6C6.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; __aza_perm=CheckPermissionCookie; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; xmlHttpRequest=true; everest_g_v2=g_surferid~XhQr8gAAFPW8EsFc; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; BIDUPSID=AD4E635AC5395DE343B6AC26FE46C873; PSTM=1580807141; BAIDUID=AD4E635AC5395DE32BF32FAB36D9518E:FG=1; _abck=DBA7C3BD927B70F835EE71C7C1AF1E5C~-1~YAAQ3tfSF1Al/c5vAQAALYYZEAP3PsVEykKgNX1Z7yztci5hDgw84xn4PLHgJApP9czi1665gl9UYf3Psj8hMewsC0VJyJ4zeYaAFnDRWBRn1V6QHFRN+AM1fdPLrdjVH03rIq8N/2uegBrLK1oqeHfP7xekO3huLcT5DQyMv7AsPC8sgBewMpaTC0r3cZ/sEsFtnKPxLxmLcpMThl0AM4o5v29/QwTJkACVqtbhSrQLE82F7JdGj0DDrhRbGu2lv036FerPg1M9d2o85I5lgfJvg6xNIT42/pvnc0U=~-1~-1~-1; sp=b75e6496-3e3c-4c58-acd1-2132ecf4d70c; _hjid=941e1bc6-293c-457e-bc91-4f5fc262dc09; _sp_id.9639=2314dd85-afcb-477b-9e82-64f1b98214d6.1580818111.2.1580906938.1580818818.be135e72-19df-41f8-acff-175cc2874545; s_vi=[CS]v1|2F1F44F18515E7A7-600006E4436E96C5[CE]; demdex=31931497131425969634038488604199850452; __cfduid=d09434e0603d6a09f49dbe43c5fe05fe11583840046; check=true; javaScript=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; screenInfo="900:1440"; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1266252805%7CMCIDTS%7C18355%7CMCMID%7C10178513813538045854600690115189903694%7CMCAAMLH-1586437468%7C11%7CMCAAMB-1586437468%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1585839869s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-18276%7CvVersion%7C4.4.1%7CMCCIDH%7C1246701492; _webvpn_key=eyJhbGciOiJIUzI1NiJ9.eyJ1c2VyIjoiMjAxNzIyNDQ5MiIsImlhdCI6MTU4NTgzNzc0NywiZXhwIjoxNTg1OTI0MTQ3fQ.FSYYdeb0JDThOpABCEyIc669m4PMT0MQ83yDv4Jt_mU; webvpn_username=2017224492%7C1585837747%7C3a0070ca91bdb64ab613250c8d988b36f1a91e80; SCSessionID=75ED331490B284627359A20C10BE9757.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=6661a66c-3b83-4176-9; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB1A91FF07EC18A2A60FFA754168065460194034E413EE92CDF5CEE181C5B57B2CABAFDF2ADE925350150D7900CAD0CA8A6D9907BD3FC99A0D55B9D018D5D381CD8; optimizelyDomainTestCookie=0.8755540362504775; optimizelyDomainTestCookie=0.10510146851569324; optimizelyDomainTestCookie=0.7587559488566185; mbox=PC#28144c226fe547c38b225004f77d29fb.22_0#1649082576|session#b01db3d7fb37480980c267be004fa76d#1585839633; s_pers=%20v8%3D1585837780294%7C1680445780294%3B%20v8_s%3DLess%2520than%25201%2520day%7C1585839580294%3B%20c19%3Dsc%253Asearch%253Adocument%2520searchform%7C1585839580343%3B%20v68%3D1585837775220%7C1585839580444%3B; optimizelyPendingLogEvents=%5B%5D; s_sess=%20s_cpc%3D0%3B%20e78%3Dtitle%2528color%2529%3B%20s_sq%3D%3B%20c21%3Ddoi%252810.1007%252Fs10462-017-9596-9%2529%3B%20e13%3Ddoi%252810.1007%252Fs10462-017-9596-9%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Asearch%25253Adocument%252520searchform%252C15%252C15%252C193%252C1055%252C193%252C1440%252C900%252C1%252CP%3B%20e41%3D1%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520searchform%252C63%252C63%252C686%252C1055%252C193%252C1440%252C900%252C1%252CP%3B',
        'Host': 'www-scopus-com-s.webvpn.nefu.edu.cn',
        # 'Referer': 'https://www-scopus-com-s.webvpn.nefu.edu.cn/search/form.uri?display=basic&zone=header&origin=',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-site',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36',
}

def myerror():
    time.sleep(1)
    try:
        browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
    except Exception:
        print('0')

def get_excel():
    file = "../sourceData/city.xls"

    data = xlrd.open_workbook(file, formatting_info=True)
    table = data.sheet_by_name('引用文献')
    papers = []
    for i in range(54, 81):
        paper = {}
        content = table.row_values(i)
        paper['标题'] = content[4]
        paper['年份'] = content[5]
        paper['DOI'] = content[14]
        print(paper['标题'])
        print(paper['DOI'])
        papers.append(paper)
    return papers


if __name__ == '__main__':
    papers = get_excel()  # 获取表格中的数据

    browser = webdriver.Chrome()
    browser.get("https://webvpn.nefu.edu.cn/")

    # 这里通过name选择器获取登录名和密码并把需要set值给放进去
    browser.find_element_by_name("username").send_keys("2017224492")
    browser.find_element_by_name("password").send_keys("040195")
    # #这一步模拟点击登录
    browser.find_element_by_class_name("login_btn").click()
    time.sleep(2)
    i = 1
    for paper in papers:
        year = paper['年份']
        doi = paper['DOI']
        if doi != '':
            url = url1 + doi + url2 + doi + url3 + doi + url4
            page_source = requests.get(url=url, headers=headers, allow_redirects=False)
            sourceID = re.findall(r'/sourceid/(.*?)\?origin=resultslist"', page_source.text, re.S)
            print(sourceID)
            if sourceID and year >= 1960:
                year = str(year)[0:4]
                sourceURL = souUrl1 + year + souUrl2 + sourceID[0] + souUrl3
                print(sourceURL)
                browser.get(sourceURL)
                if i == 1:
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"showAllPageBubble\"]/span[2]").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"selectAllMenuItem\"]/span[2]/span/ul/li[2]/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"chunkExport\"]").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"exportList\"]/li[4]/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"bibliographicalInformationCheckboxes\"]/span/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"abstractInformationCheckboxes\"]/span/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"fundInformationCheckboxes\"]/span/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"otherInformationCheckboxes\"]/span/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"otherInfoCheckboxes\"]/ul/li[4]/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"fundingCheckboxes\"]/ul/li[4]/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"exportTrigger\"]").click()

                    i = i + 1
                else:
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"showAllPageBubble\"]/span[2]").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"selectAllMenuItem\"]/span[2]/span/ul/li[2]/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"directExport\"]/span").click()

            print("---------------------")


        else:
            print("The overflow！")

    browser.close()
