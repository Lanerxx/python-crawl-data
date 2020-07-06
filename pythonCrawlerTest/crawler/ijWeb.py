import re
import urllib.request

import requests
from bs4 import BeautifulSoup
import xlrd


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
    'cookie': '__cfduid=db07216da5aea75c154e5e7236d421c3f1587562810; scopus.machineID=565E426BB5379DC87C63C542901909C3.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1587562861883r0.38711136738597673; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; xmlHttpRequest=true; SCSessionID=7F85DBF89969BE45F24C47C461F4409C.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=a6081e42-b12b-4e38-8; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB1DC8B82A34933ED66693162100B24C6AFB750D094B12FAEDD2BB6732C52682DD32CFBB76ECCEBE0946FCCE2B9E27272A3EC3592A6AC184594FE99C162F29B2C2E; check=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; javaScript=true; screenInfo="900:1440"; mbox=PC#7f4649f1b841468a940386570b585808.22_0#1651152161|session#c9dd4ba83bf7480fbfd38aecde47ce19#1587909164; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1075005958%7CMCIDTS%7C18379%7CMCMID%7C09635205066540933810671114848219444600%7CMCAAMLH-1588512162%7C11%7CMCAAMB-1588512162%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1587914562s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-18382%7CMCCIDH%7C1249006796%7CvVersion%7C4.4.1; optimizelyPendingLogEvents=%5B%5D; s_pers=%20c19%3Dsc%253Asearch%253Adocument%2520searchform%7C1587909162337%3B%20v68%3D1587907360014%7C1587909162401%3B%20v8%3D1587907397135%7C1682515397135%3B%20v8_s%3DLess%2520than%25207%2520days%7C1587909197135%3B; s_sess=%20s_sq%3D%3B%20e41%3D1%3B%20s_cpc%3D0%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Asearch%25253Adocument%252520searchform%252C75%252C75%252C789%252C1440%252C789%252C1440%252C900%252C1%252CP%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520searchform%252C75%252C75%252C789%252C1440%252C150%252C1440%252C900%252C1%252CP%3B',
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
    'cookie': '__cfduid=db07216da5aea75c154e5e7236d421c3f1587562810; scopus.machineID=565E426BB5379DC87C63C542901909C3.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1587562861883r0.38711136738597673; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; xmlHttpRequest=true; SCSessionID=7F85DBF89969BE45F24C47C461F4409C.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=a6081e42-b12b-4e38-8; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB1DC8B82A34933ED66693162100B24C6AFB750D094B12FAEDD2BB6732C52682DD32CFBB76ECCEBE0946FCCE2B9E27272A3EC3592A6AC184594FE99C162F29B2C2E; check=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; javaScript=true; screenInfo="900:1440"; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1075005958%7CMCIDTS%7C18379%7CMCMID%7C09635205066540933810671114848219444600%7CMCAAMLH-1588512162%7C11%7CMCAAMB-1588512162%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1587914562s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-18382%7CMCCIDH%7C1249006796%7CvVersion%7C4.4.1; __cfruid=64dedaa6c368b3109acc1455d54a780591c32a61-1587907545; mbox=PC#7f4649f1b841468a940386570b585808.22_0#1651152440|session#c9dd4ba83bf7480fbfd38aecde47ce19#1587909164; s_pers=%20v8%3D1587907646126%7C1682515646126%3B%20v8_s%3DLess%2520than%25207%2520days%7C1587909446126%3B%20c19%3Dsc%253Arecord%253Adocument%2520record%7C1587909446156%3B%20v68%3D1587907638814%7C1587909446207%3B; s_sess=%20s_cpc%3D0%3B%20c21%3Dtitle%2528color%2529%3B%20e13%3Dtitle%2528color%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20e78%3Dtitle%2528color%2529%3B%20s_sq%3D%3B%20e41%3D1%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Arecord%25253Adocument%252520record%252C2%252C2%252C150%252C1440%252C150%252C1440%252C900%252C1%252CP%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520searchform%252C14%252C14%252C150%252C1440%252C150%252C1440%252C900%252C1%252CP%3B',
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
    file = "../sourceData/seed.xls"
    data = xlrd.open_workbook(file,formatting_info=True)
    table = data.sheet_by_name('1')
    papers = []
    for i in range(1,16):
        paper = {}
        content = table.row_values(i)
        paper['序号'] = content[0]
        paper['标题'] = content[3]
        paper['DOI'] = content[13]
        paper['EID'] = content[43]
        papers.append(paper)
    return papers

def get_url(doi,eid):
    doi = doi.replace('/','%2f')
    url = url1 + eid + url2 + doi + url3 + doi + url4
    return url


if __name__ == '__main__':
    papers = get_excel()
    i = 1
    for paper in papers:
        doi = paper['DOI']
        eid = paper['EID']
        print(paper['序号'])
        print(paper['标题'])
        if doi != '' and eid != '':
            url = get_url(doi,eid)
            page_source = requests.get(url=url,headers=headerDetail,allow_redirects=False)
            authorUrls0 = re.findall(r'<section id="authorlist(.*?)</section>',page_source.text,re.S)
            authorUrls = re.findall(r'type="hidden"><a href="(.*?)" title="',str(authorUrls0),re.S)
            print(authorUrls)
        else:
            print("无EID")
