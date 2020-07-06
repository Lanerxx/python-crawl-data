import urllib.request
from bs4 import BeautifulSoup
import requests
import xlrd
import time
import re

YEAR = 2019

headers = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'accept-encoding': 'gzip, deflate, br',
        'accept-language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
        'cache-control': 'max-age=0',
        'cookie': '__cfduid=db07216da5aea75c154e5e7236d421c3f1587562810; scopus.machineID=565E426BB5379DC87C63C542901909C3.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1587562861883r0.38711136738597673; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; xmlHttpRequest=true; SCSessionID=1E49E1422F25B2F066355053997CDB61.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=e7885c7c-52fb-4ad2-8; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB15C6A53A420067583F4AB365B8A659D4CC7A45CE51B5BDE8F46CCE6ED53979D6BBAFDF2ADE925350150D7900CAD0CA8A6F3E74256BFBB0204C4FFD656B3875BFA; check=true; javaScript=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; optimizelyPendingLogEvents=%5B%5D; mbox=PC#7f4649f1b841468a940386570b585808.22_0#1651563765|session#116fdb0bda344ae5aa132d4c1dde49b5#1588320791; screenInfo="900:1440"; s_pers=%20v8%3D1588318968851%7C1682926968851%3B%20v8_s%3DLess%2520than%25207%2520days%7C1588320768851%3B%20c19%3Dsc%253Asearch%253Adocument%2520searchform%7C1588320768879%3B%20v68%3D1588318961677%7C1588320768991%3B; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1075005958%7CMCIDTS%7C18384%7CMCMID%7C09635205066540933810671114848219444600%7CMCAAMLH-1588923769%7C11%7CMCAAMB-1588923769%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1588326169s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-18382%7CMCCIDH%7C1249006796%7CvVersion%7C4.4.1; s_sess=%20s_sq%3D%3B%20s_ppvl%3D%3B%20e41%3D1%3B%20s_cpc%3D0%3B%20s_cc%3Dtrue%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520searchform%252C70%252C70%252C740%252C1440%252C306%252C1440%252C900%252C1%252CP%3B',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'cross-site',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36',
}

data = {
    'clusterDisplayCount': '160',
    'sot': 'cite',
    'navigatorName': '',
    'clusterCategory': '',
    'cite': 'eid',
    'refeid': '',
    'refeidnss': '',
    's': '',
    'st1': '',
    'st2': '',
    'sid': 'a9592be3f087a90dfeddc534be959fc5',
    'sdt': 'cl',
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
    'cluster': 'scopubyr,\"2017\",t,\"2016\",t',
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

clusterCategorySub = 'selectedSubjectClusterCategories'
navigatorNameSub = 'SUBJAREA'

clusterCategorySource = 'selectedSourceClusterCategories'
navigatorNameSource = 'EXACTSRCTITLE'

clusterCategoryAffi = 'selectedAffiliationClusterCategories'
navigatorNameAffi = 'AFFIL'

clusterCategoryCoun = 'selectedCountryClusterCategories'
navigatorNameCoun = 'COUNTRY_NAME'

cluster_sole1 = 'scopubyr,\"'
cluster_sole2 = '\",t'

cluster1 = 'scopubyr,\"'
# 2017
cluster2 = '\",t,\"'
# 2016
cluster3 = '\",t'

u1 = 'https://www.scopus.com/results/citedbyresults.uri?sort=plf-f&cite='
# 2-s2.0-34249309179
u2 = '&src=s&nlo=&nlr=&nls=&imp=t&sid=6e6e378fd1ca34b491c59fa079886305&sot=cite&sdt=cl&cluster=scopubyr%2C%22'
# 2008
u3 = '%22%2Ct%2C%22'
# 2007
u4 = '%22%2Ct&sl=0&origin=resultslist&zone=leftSideBar&editSaveSearch=&txGid='


u1_sole = 'https://www.scopus.com/results/citedbyresults.uri?sort=plf-f&cite='
# 2-s2.0-34249309179
u2_sole = '&src=s&nlo=&nlr=&nls=&imp=t&sid=b5cb8e30a8dc7e290b6e70f99e76daec&sot=cite&sdt=cl&cluster=scopubyr%2C%22'
# 2017
u3_sole = '%22%2Ct&sl=0&origin=resultslist&zone=leftSideBar&editSaveSearch=&txGid='


def serch(url,headers):
    req = urllib.request.Request(url=url, headers=headers)
    rsp = urllib.request.urlopen(req,timeout=20000)
    html = rsp.read().decode()
    s = BeautifulSoup(html, 'html.parser')
    return s

def get_excel():
    file = "../sourceData/a种子文献.xls"
    data = xlrd.open_workbook(file, formatting_info=True)
    table = data.sheet_by_name('1')
    papers = []
    for i in range(1,51):
        paper = {}
        content = table.row_values(i)
        paper['序号'] = content[0]
        paper['标题'] = content[3]
        paper['年份'] = content[4]
        paper['来源出版物名称'] = content[5]
        paper['DOI'] = content[13]
        paper['EID'] = content[43]
        print(paper['标题'])
        print(paper['年份'])
        print(paper['来源出版物名称'])
        print(paper['DOI'])
        print(paper['EID'])
        papers.append(paper)
    return papers

def get_subjectArea(page_source):
    subData0 = re.findall(r'<label class="checkbox-label" for=\'cat_SUBJAREA(.*?)\n</label>', page_source, re.S)
    subData = re.findall(r'class="btnText">\\n(.*?)\\n</span>', str(subData0), re.S)
    return subData

def get_country(page_source):
    counData0 = re.findall(r'<label class="checkbox-label" for=\'cat_COUNTRY(.*?)\n</label>', page_source, re.S)
    counData = re.findall(r'class="btnText">\\n(.*?)\\n</span>', str(counData0), re.S)
    return counData

def get_source(page_source):
    subData0 = re.findall(r'<label class="checkbox-label" for=\'cat_EXACTSRCTITLE(.*?)\n</label>', page_source, re.S)
    subData = re.findall(r'class="btnText">\\n(.*?)\\n</span>', str(subData0), re.S)
    return subData

if __name__ == '__main__':

    papers = get_excel()
    fEarlyCitation = open('../data/earlyCit.txt', 'w')
    fEarlySub = open('../data/earlySub.txt', 'w')
    fEarlySource = open('../data/earlySource.txt', 'w')
    fEarlyAffi = open('../data/earlyAffi.txt', 'w')
    fEarlyCoun = open('../data/earlyCoun.txt', 'w')

    for paper in papers:
        eid = paper['EID']
        year = paper['年份']

        # ====================================1=====================================

        if eid != '':
            if year >= YEAR-1:
                url = u1_sole + eid + u2_sole + str(YEAR-1)[0:4] + u3_sole
            else:
                url = u1 + eid + u2 + str(year+2)[0:4] + u3 + str(year+1)[0:4] + u4
            print(url)
            page_source = requests.get(url=url, headers=headers, allow_redirects=False)
            # -----------------引用-------------------
            count = re.findall(r'<span class="resultsCount">\n(.*?)\n</span>', page_source.text, re.S)
            print("earlyCites:")
            if count:
                print(count[0])
                print(count[0], file=fEarlyCitation, flush=True)
            else:
                print('', file=fEarlyCitation, flush=True)
            print("\n")
            # -----------------学科-------------------

            subjectArea = get_subjectArea(page_source.text)
            earlySub = len(subjectArea)
            if earlySub >= 10:
                dataSub = data
                dataSub['navigatorName'] = navigatorNameSub
                dataSub['clusterCategory'] = clusterCategorySub
                dataSub['cite'] = eid
                dataSub['cluster'] = cluster1 + str(year+2)[0:4] + cluster2 + str(year+1)[0:4] + cluster3
                if year >= YEAR-1:
                    dataSub['cluster'] = cluster_sole1 + str(YEAR-1)[0:4] + cluster_sole2

                rep = requests.post(
                    url='https://www.scopus.com/standard/viewMore.uri', data=dataSub,
                    headers=headers)
                earlySub0 = re.findall(r'btnText\\">(.*?)<', rep.text, re.S)
                subjectArea = earlySub0
                earlySub = len(earlySub0) / 2
            print("Subject area:")
            print(subjectArea)
            print(earlySub)
            print("\n")
            print(earlySub, file=fEarlySub, flush=True)
            time.sleep(1)

            # -----------------来源-------------------


            source = get_source(page_source.text)
            earlySource = len(source)

            if earlySource >= 10:
                dataSource = data
                dataSource['navigatorName'] = navigatorNameSource
                dataSource['clusterCategory'] = clusterCategorySource
                dataSource['cite'] = eid
                dataSource['cluster'] = cluster1 + str(year+2)[0:4] + cluster2 + str(year+1)[0:4] + cluster3
                if year >= YEAR-1:
                    dataSource['cluster'] = cluster_sole1 + str(YEAR-1)[0:4] + cluster_sole2

                rep = requests.post(
                    url='https://www.scopus.com/standard/viewMore.uri', data=dataSource,
                    headers=headers)
                earlySource0 = re.findall(r'btnText\\">(.*?)<', rep.text, re.S)
                source = earlySource0
                earlySource = len(earlySource0) / 2
            print("Source:")
            print(source)
            print(earlySource)
            print("\n")
            print(earlySource, file=fEarlySource, flush=True)
            time.sleep(1)

            # -----------------机构-------------------
            dataAffi = data
            dataAffi['navigatorName'] = navigatorNameAffi
            dataAffi['clusterCategory'] = clusterCategoryAffi
            dataAffi['cite'] = eid
            dataAffi['cluster'] = cluster1 + str(year + 2)[0:4] + cluster2 + str(year + 1)[0:4] + cluster3
            if year >= YEAR-1:
                dataAffi['cluster'] = cluster_sole1 + str(YEAR-1)[0:4] + cluster_sole2

            rep = requests.post(
                url='https://www.scopus.com/standard/viewMore.uri', data=dataAffi,
                headers=headers)
            earlyAffi0 = re.findall(r'btnText\\">(.*?)<', rep.text, re.S)
            affilication = earlyAffi0
            earlyAffi = len(earlyAffi0) / 2
            print("Affilication:")
            print(affilication)
            print(earlyAffi)
            print("\n")
            print(earlyAffi, file=fEarlyAffi, flush=True)
            time.sleep(1)

            # -----------------国家-------------------

            country = get_country(page_source.text)
            earlyCoun = len(country)

            if earlyCoun == 10:
                dataCoun = data
                dataCoun['navigatorName'] = navigatorNameCoun
                dataCoun['clusterCategory'] = clusterCategoryCoun
                dataCoun['cite'] = eid
                dataCoun['cluster'] = cluster1 + str(year + 2)[0:4] + cluster2 + str(year + 1)[0:4] + cluster3
                if year >= YEAR-1:
                    dataCoun['cluster'] = cluster_sole1 + str(YEAR-1)[0:4] + cluster_sole2

                rep = requests.post(
                    url='https://www.scopus.com/standard/viewMore.uri', data=dataCoun,
                    headers=headers)
                earlyCoun0 = re.findall(r'btnText\\">(.*?)<', rep.text, re.S)
                country = earlyCoun0
                earlyCoun = len(earlyCoun0) / 2

            print("Country:")
            print(country)
            print(earlyCoun)
            print("\n")
            print(earlyCoun, file=fEarlyCoun, flush=True)
            time.sleep(1)

        else:
            print('', file=fEarlyCitation, flush=True)
            print('', file=fEarlySub, flush=True)
            print('', file=fEarlySource, flush=True)
            print('', file=fEarlyAffi, flush=True)
            print('', file=fEarlyCoun, flush=True)



        print("=======================================================================")



