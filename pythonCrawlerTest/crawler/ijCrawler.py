import re
import urllib.request

import requests
from bs4 import BeautifulSoup
import xlrd
from selenium import webdriver
import time

url1 = 'https://www.scopus.com/record/display.uri?eid='
# 2-s2.0-85049078383
url2 = '&origin=resultslist&sort=plf-f&src=s&st1='
# 10.1007%2fs11192-018-2824-5
url3 = '&st2=&sid=050580c4789bac0b82a163ff23860139&sot=b&sdt=b&sl=30&s=DOI%28'
# 10.1007%2fs11192-018-2824-5
url4 = '%29&relpos=0&citeCnt=25&searchTerm='

headers = {
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'accept-encoding': 'gzip, deflate, br',
    'accept-language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
    'cache-control': 'max-age=0',
    'cookie': 'optimizelyEndUserId=oeu1577535319518r0.069338915535341; optimizelyBuckets=%7B%7D; xmlHttpRequest=true; scopus.machineID=9349BCB0F58400EBEEE3542E261FC722.wsnAw8kcdt7IPYLO0V48gA; check=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; javaScript=true; screenInfo="900:1440"; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; __cfruid=1ae81abd72848f395064971cef28d4ed34493be3-1587371929; __cfduid=d5af55cba1dad743d7bfa36874b02bce61587373707; SCSessionID=C48B88B751ADF007E858F70F914081AE.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=4b8de978-e812-4540-b; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB18C9834EB8F99A8CE068326372565D2998B665F7C2B0A999498AD3C5B447E80B5BAFDF2ADE925350150D7900CAD0CA8A68C5CE404F394CA075DBDFAE6664FE64F; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1075005958%7CMCIDTS%7C18373%7CMCMID%7C75774528907604783432259580220259513125%7CMCAAMLH-1587992139%7C11%7CMCAAMB-1587992139%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1587394540s%7CNONE%7CMCAID%7CNONE%7CMCCIDH%7C1249006796%7CvVersion%7C4.4.1; optimizelyPendingLogEvents=%5B%5D; mbox=PC#4f355c2658574c81ad8213018b82db66.22_0#1650633588|session#5dda94991b1842a19c7910a6ebbfc065#1587389194; s_pers=%20v8%3D1587388790778%7C1681996790778%3B%20v8_s%3DLess%2520than%25201%2520day%7C1587390590778%3B%20c19%3Dsc%253Asearch%253Adocument%2520searchform%7C1587390590821%3B%20v68%3D1587388785515%7C1587390590899%3B; s_sess=%20s_cpc%3D0%3B%20e78%3Ddoi%252810.1007%252Fs11192-018-2695-9%2529%3B%20c21%3Ddoi%252810.1007%252Fs11192-018-2711-0%2529%3B%20e13%3Ddoi%252810.1007%252Fs11192-018-2711-0%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20s_sq%3D%3B%20e41%3D1%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Arecord%25253Adocument%252520record%252C4%252C4%252C432%252C1440%252C432%252C1440%252C900%252C1%252CP%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520searchform%252C20%252C20%252C207%252C1440%252C207%252C1440%252C900%252C1%252CP%3B',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'same-origin',
    'sec-fetch-user': '?1',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36',
}

headerDetail = {
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'accept-encoding': 'gzip, deflate, br',
    'accept-language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
    'cache-control': 'max-age=0',
    'cookie': 'optimizelyEndUserId=oeu1577535319518r0.069338915535341; optimizelyBuckets=%7B%7D; xmlHttpRequest=true; scopus.machineID=9349BCB0F58400EBEEE3542E261FC722.wsnAw8kcdt7IPYLO0V48gA; check=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; javaScript=true; screenInfo="900:1440"; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; __cfruid=1ae81abd72848f395064971cef28d4ed34493be3-1587371929; __cfduid=d5af55cba1dad743d7bfa36874b02bce61587373707; SCSessionID=C48B88B751ADF007E858F70F914081AE.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=4b8de978-e812-4540-b; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB18C9834EB8F99A8CE068326372565D2998B665F7C2B0A999498AD3C5B447E80B5BAFDF2ADE925350150D7900CAD0CA8A68C5CE404F394CA075DBDFAE6664FE64F; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1075005958%7CMCIDTS%7C18373%7CMCMID%7C75774528907604783432259580220259513125%7CMCAAMLH-1587992139%7C11%7CMCAAMB-1587992139%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1587394540s%7CNONE%7CMCAID%7CNONE%7CMCCIDH%7C1249006796%7CvVersion%7C4.4.1; optimizelyPendingLogEvents=%5B%5D; s_pers=%20c19%3Dsc%253Arecord%253Aauthor%2520details%7C1587389528838%3B%20v68%3D1587387725538%7C1587389528936%3B%20v8%3D1587388766762%7C1681996766762%3B%20v8_s%3DLess%2520than%25201%2520day%7C1587390566762%3B; s_sess=%20s_cpc%3D0%3B%20e78%3Ddoi%252810.1007%252Fs11192-018-2695-9%2529%3B%20c21%3Ddoi%252810.1007%252Fs11192-018-2711-0%2529%3B%20e13%3Ddoi%252810.1007%252Fs11192-018-2711-0%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20e41%3D1%3B%20s_cc%3Dtrue%3B%20s_sq%3Delsevier-sc-prod%25252Celsevier-global-prod%253D%252526pid%25253Dsc%2525253Arecord%2525253Aauthor%25252520details%252526pidt%25253D1%252526oid%25253Dhttps%2525253A%2525252F%2525252Fwww.scopus.com%2525252Fhome.uri%2525253Fzone%2525253Dheader%25252526origin%2525253DAuthorProfile%252526ot%25253DA%3B%20s_ppvl%3Dsc%25253Arecord%25253Aauthor%252520details%252C20%252C20%252C432%252C1440%252C432%252C1440%252C900%252C1%252CP%3B%20s_ppv%3Dsc%25253Arecord%25253Adocument%252520record%252C4%252C4%252C432%252C1440%252C432%252C1440%252C900%252C1%252CP%3B; mbox=PC#4f355c2658574c81ad8213018b82db66.22_0#1650633578|session#5dda94991b1842a19c7910a6ebbfc065#1587389194',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'none',
    'sec-fetch-user': '?1',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36',

}

def serch(url,headers):
    req = urllib.request.Request(url=url,headers=headers)
    rsp = urllib.request.urlopen(req,timeout=20000)
    html = rsp.read().decode()
    s = BeautifulSoup(html,'html.parser')
    return s

def get_excel():
    file = "../sourceData/种子文献.xls"
    data = xlrd.open_workbook(file,formatting_info=True)
    table = data.sheet_by_name('scopus')
    papers = []
    for i in range(1,2):
        paper = {}
        content = table.row_values(i)
        paper['标题'] = content[1]
        paper['DOI'] = content[15]
        paper['EID'] = content[14]
        print(paper['标题'])
        print(paper['DOI'])
        print(paper['EID'])
        papers.append(paper)
    return papers

def get_url(doi,eid):
    doi = doi.replace('/','%2f')
    url = url1 + eid + url2 + doi + url3 + doi + url4
    return url

def myerror():
    time.sleep(1)
    try:
        browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
    except Exception:
        print('0')

def get_exPaper(authorUrl,eid):
    print(eid)
    n = 0
    m = 0
    page_source = requests.get(url=authorUrl,headers=headers,allow_redirects=False)
    print(page_source)
    paperEids0 = re.findall(r'<td data-type="docTitle">(.*?)</td>',page_source.text,re.S)
    paperEids = re.findall(r'eid=(.*?)&amp',str(paperEids0),re.S)
    print(paperEids)
    for paperEid in paperEids:
        if eid == paperEid:
            n = n + 1
            m = 1
            break
        else:
            n = n + 1

    if m != 0:
        xPath = '//*[@id=\"resultDataRow0\"]/td[' + n + ']/a'
        browser.find_element_by_xpath(xPath).click()


if __name__ == '__main__':
    papers = get_excel()
    browser = webdriver.Chrome()
    browser.get(
        "https://id.elsevier.com/as/authorization.oauth2?platSite=SC%2Fscopus&ui_locales=en-US&scope=openid+profile+email+els_auth_info+els_analytics_info+urn%3Acom%3Aelsevier%3Aidp%3Apolicy%3Aproduct%3Aindv_identity&response_type=code&redirect_uri=https%3A%2F%2Fwww.scopus.com%2Fauthredirect.uri%3FtxGid%3Df5a2c868e544620e3dd9c4a9bfdde55a&state=userLogin&authType=SINGLE_SIGN_IN&prompt=login&client_id=SCOPUS")
    browser.find_element_by_xpath("//*[@id=\"bdd-elsSecondaryBtn\"]/span/span").click()
    browser.find_element_by_xpath("//*[@id=\"bdd-email\"]").click()
    browser.find_element_by_name("els_institution").send_keys("l.y.x.peng@nefu.edu.cn")
    time.sleep(1)
    browser.find_element_by_xpath("//*[@id=\"bdd-els-searchBtn\"]").click()
    time.sleep(1)
    browser.find_element_by_name("password").send_keys("yndj0401")
    browser.find_element_by_xpath("//*[@id=\"bdd-elsPrimaryBtn\"]").click()
    time.sleep(2)
    i = 1
    for paper in papers:
        doi = paper['DOI']
        eid = paper['EID']

        if doi != '' and eid != '':
            url = get_url(doi,eid)
            page_source = requests.get(url=url,headers=headerDetail,allow_redirects=False)
            authorUrls0 = re.findall(r'<section id="authorlist(.*?)</section>',page_source.text,re.S)
            authorUrls = re.findall(r'type="hidden"><a href="(.*?)" title="',str(authorUrls0),re.S)
            for authorUrl in authorUrls:
                print(authorUrl)
                browser.get(authorUrl)
                myerror()
                browser.find_element_by_xpath("//*[@id=\"docTabTitleBar\"]/a/span[1]").click()

                # 限制年限，排除2020年
                myerror()
                try:
                    browser.find_element_by_xpath("//*[@id=\"li_2020\"]/label").click()
                    browser.find_element_by_xpath("//*[@id=\"RefineResults\"]/div[1]/div[2]/ul/li[2]/input").click()
                except Exception:
                    print("no 2020")

                # 选择文献
                time.sleep(8)
                myerror()
                browser.find_element_by_xpath("//*[@id=\"selectAllCheck\"]/label\"]/span/label").click()
                # 排除种子文献
                get_exPaper(authorUrl,eid)

                if i == 1:

                    # 按需导出
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"export_results\"]/span").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"exportList\"]/li[4]/label").click()

                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"citationGroupCheckboxes\"]/span/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"bibliographicalInformationCheckboxes\"]/span/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"abstractInformationCheckboxes\"]/span/label").click()

                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"fundingCheckboxes\"]/ul/li[1]/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"otherInformationCheckboxes\"]/span/label").click()
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"otherInfoCheckboxes\"]/ul/li[4]/label").click()

                    time.sleep(2)
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"exportTrigger\"]/span").click()
                    time.sleep(4)

                    i = i + 1
                else:
                    myerror()
                    browser.find_element_by_xpath("//*[@id=\"directExport\"]/span").click()

        else:
            print("无EID")

    browser.close()
