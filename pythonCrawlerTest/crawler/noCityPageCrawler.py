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
        'Cookie': 'UM_distinctid=16f03aeda951f7-0e445c640904db-12316b5a-13c680-16f03aeda96100; scopus.machineID=42A3D9C620FA39CD631B1EEC603CF6C6.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; __aza_perm=CheckPermissionCookie; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; uuid2=341871610905020489; xmlHttpRequest=true; everest_g_v2=g_surferid~XhQr8gAAFPW8EsFc; dpm=10213765694604764764599412911321458779; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; BIDUPSID=AD4E635AC5395DE343B6AC26FE46C873; PSTM=1580807141; BAIDUID=AD4E635AC5395DE32BF32FAB36D9518E:FG=1; _abck=DBA7C3BD927B70F835EE71C7C1AF1E5C~-1~YAAQ3tfSF1Al/c5vAQAALYYZEAP3PsVEykKgNX1Z7yztci5hDgw84xn4PLHgJApP9czi1665gl9UYf3Psj8hMewsC0VJyJ4zeYaAFnDRWBRn1V6QHFRN+AM1fdPLrdjVH03rIq8N/2uegBrLK1oqeHfP7xekO3huLcT5DQyMv7AsPC8sgBewMpaTC0r3cZ/sEsFtnKPxLxmLcpMThl0AM4o5v29/QwTJkACVqtbhSrQLE82F7JdGj0DDrhRbGu2lv036FerPg1M9d2o85I5lgfJvg6xNIT42/pvnc0U=~-1~-1~-1; sp=b75e6496-3e3c-4c58-acd1-2132ecf4d70c; _hjid=941e1bc6-293c-457e-bc91-4f5fc262dc09; _sp_id.9639=2314dd85-afcb-477b-9e82-64f1b98214d6.1580818111.2.1580906938.1580818818.be135e72-19df-41f8-acff-175cc2874545; s_vi=[CS]v1|2F1F44F18515E7A7-600006E4436E96C5[CE]; demdex=31931497131425969634038488604199850452; __cfduid=d09434e0603d6a09f49dbe43c5fe05fe11583840046; javaScript=true; check=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; ev_sync_dd=20200406; screenInfo="900:1440"; JSESSIONID=a181578538850a54; _webvpn_key=eyJhbGciOiJIUzI1NiJ9.eyJ1c2VyIjoiMjAxNzIyNDQ5MiIsImlhdCI6MTU4NjE3MDAxMSwiZXhwIjoxNTg2MjU2NDExfQ.Vx3xKeZ_NlSddpeC6xk_qaBl-yZPZWAaiXnXiqbnhKs; webvpn_username=2017224492%7C1586170011%7Cfb575c30866a16f9a069153d067e0dcf3ae4f799; SCSessionID=8C1371D3F653688F2E2C3143B0E40EB9.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=6454be9b-992a-4b9d-b; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB1F373C6B7FF9169B6167286DDBBB33C212E68CC500EE6499F2A4BDE26DBF07937A31AAC5A6BDE3E4B4DACF34F3854CEEB5594DF804FDBE0B9A48B096EB60A7957; optimizelyDomainTestCookie=0.8551787030572016; optimizelyDomainTestCookie=0.42173512961572523; optimizelyDomainTestCookie=0.7734683114548839; mbox=PC#28144c226fe547c38b225004f77d29fb.22_0#1649414832|session#3ff71f9ad99049ddb332a7a22de66245#1586171891; anj=dTM7k!M4/8F7/.XF\']wIg2IliwTJCJ!]G(^Lgm@*EG9EUlb:kv%v4VB%w+_m!?#^NewL+B; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1266252805%7CMCIDTS%7C18359%7CMCMID%7C10178513813538045854600690115189903694%7CMCAAMLH-1586774839%7C11%7CMCAAMB-1586774839%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1586177239s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-18276%7CvVersion%7C4.4.1%7CMCCIDH%7C1246701492; s_pers=%20v8%3D1586170039606%7C1680778039606%3B%20v8_s%3DLess%2520than%25201%2520day%7C1586171839606%3B%20c19%3Dsc%253Asearch%253Adocument%2520searchform%7C1586171839656%3B%20v68%3D1586170032982%7C1586171839901%3B; optimizelyPendingLogEvents=%5B%5D; s_sess=%20s_cpc%3D0%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Asearch%25253Adocument%252520searchform%252C25%252C25%252C267%252C1440%252C267%252C1440%252C900%252C1%252CP%3B%20e41%3D1%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520searchform%252C75%252C75%252C789%252C1440%252C267%252C1440%252C900%252C1%252CP%3B',
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
    file = "../sourceData/a1-10.xls"

    data = xlrd.open_workbook(file, formatting_info=True)
    table = data.sheet_by_name('引用文献')
    papers = []
    for i in range(317, 378):
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
