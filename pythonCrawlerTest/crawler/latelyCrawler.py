import requests
import io
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
        'Cookie': 'UM_distinctid=16f03aeda951f7-0e445c640904db-12316b5a-13c680-16f03aeda96100; scopus.machineID=42A3D9C620FA39CD631B1EEC603CF6C6.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; __aza_perm=CheckPermissionCookie; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; demdex=10213765694604764764599412911321458779; uuid2=341871610905020489; xmlHttpRequest=true; everest_g_v2=g_surferid~XhQr8gAAFPW8EsFc; dpm=10213765694604764764599412911321458779; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; BIDUPSID=AD4E635AC5395DE343B6AC26FE46C873; PSTM=1580807141; BAIDUID=AD4E635AC5395DE32BF32FAB36D9518E:FG=1; _abck=DBA7C3BD927B70F835EE71C7C1AF1E5C~-1~YAAQ3tfSF1Al/c5vAQAALYYZEAP3PsVEykKgNX1Z7yztci5hDgw84xn4PLHgJApP9czi1665gl9UYf3Psj8hMewsC0VJyJ4zeYaAFnDRWBRn1V6QHFRN+AM1fdPLrdjVH03rIq8N/2uegBrLK1oqeHfP7xekO3huLcT5DQyMv7AsPC8sgBewMpaTC0r3cZ/sEsFtnKPxLxmLcpMThl0AM4o5v29/QwTJkACVqtbhSrQLE82F7JdGj0DDrhRbGu2lv036FerPg1M9d2o85I5lgfJvg6xNIT42/pvnc0U=~-1~-1~-1; sp=b75e6496-3e3c-4c58-acd1-2132ecf4d70c; _hjid=941e1bc6-293c-457e-bc91-4f5fc262dc09; _sp_id.9639=2314dd85-afcb-477b-9e82-64f1b98214d6.1580818111.2.1580906938.1580818818.be135e72-19df-41f8-acff-175cc2874545; __cfduid=d5f4fadf6ae3c94893e08d9ef00c501281581146065; s_vi=[CS]v1|2F1F44F18515E7A7-600006E4436E96C5[CE]; _webvpn_key=eyJhbGciOiJIUzI1NiJ9.eyJ1c2VyIjoiMjAxNzIyNDQ5MiIsImlhdCI6MTU4Mjk2MDk3MiwiZXhwIjoxNTgzMDQ3MzcyfQ.5AWwT_2wKdN4T2w5Hf_3Bxi1HOv-qf5UkdPuhV183mw; webvpn_username=2017224492%7C1582960972%7Cb90c7ce464fde268b5bded60252a7a5de9687736; SCSessionID=296DDD7587332A8344FC2F06F216A685.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=3d7faf0b-10d6-4c1e-b; check=true; javaScript=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1266252805%7CMCIDTS%7C18321%7CMCMID%7C10178513813538045854600690115189903694%7CMCAAMLH-1583565834%7C11%7CMCAAMB-1583565834%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1582968235s%7CNONE%7CMCAID%7CNONE%7CMCCIDH%7C1246701492%7CMCSYNCSOP%7C411-18276%7CvVersion%7C4.4.1; ev_sync_dd=20200229; screenInfo="900:1440"; JSESSIONID=23657e62430e326; __cfruid=aab26e97fe7a0adbe1d134057166bc7f856e38be-1582961479; anj=dTM7k!M4/8F7/.XF\']wIg2IliwTJCJ!s8K%KkmY]\'s%YY9sk@3@\'u0e4(vAN!#E/r; optimizelyDomainTestCookie=0.6945536417589959; optimizelyDomainTestCookie=0.5862335756343324; optimizelyDomainTestCookie=0.038820288226025346; mbox=PC#28144c226fe547c38b225004f77d29fb.22_0#1646206357|session#1285e4a8527845549f5787d3b37607e3#1582962884; s_pers=%20v8%3D1582961562578%7C1677569562578%3B%20v8_s%3DLess%2520than%25201%2520day%7C1582963362578%3B%20c19%3Dsc%253Asearch%253Adocument%2520results%7C1582963362868%3B%20v68%3D1582961555813%7C1582963363113%3B; s_sess=%20s_cpc%3D0%3B%20c21%3Ddoi%252810.1257%252Fjep.10.3.153%2529%3B%20e13%3Ddoi%252810.1257%252Fjep.10.3.153%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20e41%3D1%3B%20s_sq%3D%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Asearch%25253Adocument%252520results%252C56%252C42%252C1084%252C1440%252C789%252C1440%252C900%252C1%252CP%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520results%252C56%252C38%252C1084%252C1440%252C248%252C1440%252C900%252C1%252CP%3B',
        'Host': 'www-scopus-com-s.webvpn.nefu.edu.cn',
        # 'Referer': 'https://www-scopus-com-s.webvpn.nefu.edu.cn/search/form.uri?display=basic&zone=header&origin=',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.122 Safari/537.36',
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

cluster_17 = 'scopubyr,\"2017\",t'

def get_excel():
    file = "./data.xls"
    data = xlrd.open_workbook(file, formatting_info=True)
    table = data.sheet_by_name('引用文献')
    papers = []
    for i in range(1,4):
        paper = {}
        content = table.row_values(i)
        paper['论文序号'] = content[1]
        paper['作者'] = content[2]
        paper['作者ID'] = content[3]
        paper['标题'] = content[4]
        paper['年份'] = content[5]
        paper['来源出版物名称'] = content[6]
        paper['DOI'] = content[8]
        paper['EID'] = content[13]
        print(paper['EID'])
        papers.append(paper)
    return papers

if __name__ == '__main__':

    papers = get_excel()
    fLatelySub = open('./latelySub.txt', 'w')
    fLatelySource = open('./latelySource.txt', 'w')
    fLatelyAffi = open('./latelyAffi.txt', 'w')
    fLatelyCoun = open('./latelyCoun.txt', 'w')


    for paper in papers:
        eid = paper['EID']
        year = paper['年份']



        if eid != '':

                # -----------------学科-------------------

                dataSub = data
                dataSub['navigatorName'] = navigatorNameSub
                dataSub['clusterCategory'] = clusterCategorySub
                dataSub['cite'] = eid
                if year == 2017:
                    dataSub['cluster'] = cluster_17

                rep = requests.post(
                    url='https://www-scopus-com-s.webvpn.nefu.edu.cn/standard/viewMore.uri', data=dataSub,
                    headers=headers)
                latelySub0 = re.findall(r'btnText\\">(.*?)<', rep.text, re.S)
                print("\nlatelySub:")
                latelySub = len(latelySub0) / 2
                print(latelySub)
                print(latelySub, file=fLatelySub, flush=True)
                time.sleep(1)
                # -----------------来源-------------------

                dataSource = data
                dataSource['navigatorName'] = navigatorNameSource
                dataSource['clusterCategory'] = clusterCategorySource
                dataSource['cite'] = eid
                if year == 2017:
                    dataSource['cluster'] = cluster_17

                rep = requests.post(
                    url='https://www-scopus-com-s.webvpn.nefu.edu.cn/standard/viewMore.uri', data=dataSource,
                    headers=headers)
                latelySource0 = re.findall(r'btnText\\">(.*?)<', rep.text, re.S)
                print("\nlatelySource:")
                latelySource = len(latelySource0) / 2
                print(latelySource)
                print(latelySource, file=fLatelySource, flush=True)
                time.sleep(1)

                # -----------------机构-------------------
                dataAffi = data
                dataAffi['navigatorName'] = navigatorNameAffi
                dataAffi['clusterCategory'] = clusterCategoryAffi
                dataAffi['cite'] = eid
                if year == 2017:
                    dataAffi['cluster'] = cluster_17

                rep = requests.post(
                    url='https://www-scopus-com-s.webvpn.nefu.edu.cn/standard/viewMore.uri', data=dataAffi,
                    headers=headers)
                latelyAffi0 = re.findall(r'btnText\\">(.*?)<', rep.text, re.S)
                print("\nlatelyAffi:")
                latelyAffi = len(latelyAffi0) / 2
                print(latelyAffi)
                print(latelyAffi, file=fLatelyAffi, flush=True)
                time.sleep(1)
                # -----------------国家-------------------
                dataCoun = data
                dataCoun['navigatorName'] = navigatorNameCoun
                dataCoun['clusterCategory'] = clusterCategoryCoun
                dataCoun['cite'] = eid
                if year == 2017:
                    dataCoun['cluster'] = cluster_17

                rep = requests.post(
                    url='https://www-scopus-com-s.webvpn.nefu.edu.cn/standard/viewMore.uri', data=dataCoun,
                    headers=headers)
                latelyCoun0 = re.findall(r'btnText\\">(.*?)<', rep.text, re.S)
                print("\nlatelyCoun:")
                latelyCoun = len(latelyCoun0) / 2
                print(latelyCoun)
                print(latelyCoun, file=fLatelyCoun, flush=True)
                time.sleep(1)

        else:
            print('', file=fLatelySub, flush=True)
            print('', file=fLatelySource, flush=True)
            print('', file=fLatelyAffi, flush=True)
            print('', file=fLatelyCoun, flush=True)


