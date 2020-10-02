import requests
import io
import xlrd
import time
import re
import sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='gb18030')

headers = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'accept-encoding': 'gzip, deflate, br',
        'accept-language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
        'cache-control': 'max-age=0',
        'cookie': '__cfduid=db07216da5aea75c154e5e7236d421c3f1587562810; scopus.machineID=565E426BB5379DC87C63C542901909C3.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1587562861883r0.38711136738597673; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; xmlHttpRequest=true; javaScript=true; check=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; screenInfo="900:1440"; __cfruid=967796fd4e05e623ae8db44a22ed07b3b0dc7f48-1588552488; SCSessionID=25F0F60889FFD9926BA4B6A885077D77.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=827c5ecc-ad3b-48c2-9; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB16F22DB7CAAB0FAC6BB903011EC6294D39AF88E344913E7DD4EAA9B9BC40612B210BA32070D9964CEACBAE7C5777723B713F15DDA1EB73EFD5933C091CC9E5C1F; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1075005958%7CMCIDTS%7C18388%7CMCMID%7C09635205066540933810671114848219444600%7CMCAAMLH-1589242662%7C11%7CMCAAMB-1589242662%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1588645062s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-18382%7CMCCIDH%7C1249006796%7CvVersion%7C4.4.1; mbox=PC#7f4649f1b841468a940386570b585808.22_0#1651883824|session#b61756272bbe4f9997803ac98800cb14#1588639722; s_pers=%20v8%3D1588639026552%7C1683247026552%3B%20v8_s%3DLess%2520than%25201%2520day%7C1588640826552%3B%20c19%3Dsc%253Asearch%253Adocument%2520searchform%7C1588640826589%3B%20v68%3D1588639022638%7C1588640826703%3B; optimizelyPendingLogEvents=%5B%5D; s_sess=%20s_cpc%3D0%3B%20c21%3Ddoi%252810.1109%252Ficcv.2011.6126553%2529%3B%20e13%3Ddoi%252810.1109%252Ficcv.2011.6126553%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20e78%3Ddoi%252810.1109%252Ficcv.2011.6126553%2529%3B%20s_sq%3D%3B%20e41%3D1%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Asearch%25253Adocument%252520searchform%252C75%252C75%252C789%252C1440%252C789%252C1440%252C900%252C1%252CP%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520searchform%252C75%252C75%252C789%252C1440%252C150%252C1440%252C900%252C1%252CP%3B',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'cross-site',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36',
}
headerDetail = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'accept-encoding': 'gzip, deflate, br',
        'accept-language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
        'cache-control': 'max-age=0',
        'cookie': '__cfduid=db07216da5aea75c154e5e7236d421c3f1587562810; scopus.machineID=565E426BB5379DC87C63C542901909C3.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1587562861883r0.38711136738597673; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; xmlHttpRequest=true; javaScript=true; check=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; screenInfo="900:1440"; SCSessionID=25F0F60889FFD9926BA4B6A885077D77.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=827c5ecc-ad3b-48c2-9; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB16F22DB7CAAB0FAC6BB903011EC6294D39AF88E344913E7DD4EAA9B9BC40612B210BA32070D9964CEACBAE7C5777723B713F15DDA1EB73EFD5933C091CC9E5C1F; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1075005958%7CMCIDTS%7C18388%7CMCMID%7C09635205066540933810671114848219444600%7CMCAAMLH-1589242662%7C11%7CMCAAMB-1589242662%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1588645062s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-18382%7CMCCIDH%7C1249006796%7CvVersion%7C4.4.1; __cfruid=8378526056c7df9d51f938106d80f0b285611fd5-1588639130; optimizelyPendingLogEvents=%5B%5D; mbox=PC#7f4649f1b841468a940386570b585808.22_0#1651883949|session#b61756272bbe4f9997803ac98800cb14#1588639722; s_pers=%20v8%3D1588639150947%7C1683247150947%3B%20v8_s%3DLess%2520than%25201%2520day%7C1588640950947%3B%20c19%3Dsc%253Arecord%253Adocument%2520record%7C1588640950981%3B%20v68%3D1588639148051%7C1588640951046%3B; s_sess=%20s_cpc%3D0%3B%20c21%3Dtitle-abs-key%2528color%2529%3B%20e13%3Dtitle-abs-key%2528color%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20e78%3Dtitle-abs-key%2528color%2529%3B%20s_sq%3D%3B%20e41%3D1%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Arecord%25253Adocument%252520record%252C15%252C15%252C328%252C1440%252C328%252C1440%252C900%252C1%252CP%3B%20s_ppv%3Dsc%25253Arecord%25253Adocument%252520record%252C15%252C15%252C328%252C1440%252C328%252C1440%252C900%252C1%252CP%3B',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'same-origin',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36',
}


url1 = 'https://www.scopus.com/results/results.uri?numberOfFields=0&src=s&clickedLink=&edit=&editSaveSearch=&origin=searchbasic&authorTab=&affiliationTab=&advancedTab=&scint=1&menu=search&tablin=&searchterm1='
url2 = '&field1=DOI&dateType=Publication_Date_Type&yearFrom=Before+1960&yearTo=Present&loadDate=7&documenttype=All&accessTypes=All&resetFormLink=&st1='
url3 = '&st2=&sot=b&sdt=b&sl=22&s=DOI%28'
url4 = '%29&sid=04d9016932b494f613c131f956db3e87&searchId=04d9016932b494f613c131f956db3e87&txGid=cd4965c8585c87e1b28188380fa2685e&sort=plf-f&originationType=b&rr='

citationturl1 = 'https://www.scopus.com/search/submit/citedby.uri?eid='
citationturl2 = '&src=s&origin=recordpage'

snipurl = 'https://www.scopus.com/api/rest/sources/'

detail1 = 'https://www.scopus.com/record/display.uri?eid='
# 2-s2.0-84944149175
detail2 = '&origin=resultslist&sort=plf-f&src=s&st1='
# 10.1007
detail3 = '%2f'
# s11192-014-1269-8
detail4 = '&st2=&sid=bf43c8536b4c20bfb146383e8ae724aa&sot=b&sdt=b&sl=30&s=DOI%2810.'
# 1007
detail5 = '%2f'
# s11192-014-1269-8
detail6 = '%29&relpos=0&citeCnt=46&searchTerm='

snip1 = 'https://www.scopus.com/results/results.uri?numberOfFields=0&src=s&clickedLink=&edit=&editSaveSearch=&origin=searchbasic&authorTab=&affiliationTab=&advancedTab=&scint=1&menu=search&tablin=&searchterm1='
# 10.1016%2Fj.joi.2009.11.002
snip2 = '&field1=DOI&dateType=Publication_Date_Type&yearFrom=Before+1960&yearTo=Present&loadDate=7&documenttype=All&accessTypes=All&resetFormLink=&st1='
# 10.1016%2Fj.joi.2009.11.002
snip3 = '&st2=&sot=b&sdt=b&sl=30&s=DOI%28'
# 10.1016%2Fj.joi.2009.11.002
snip4 = '%29&sid=f90ade8a461ec3d0e4486e0fb8eb8e48&searchId=f90ade8a461ec3d0e4486e0fb8eb8e48&txGid=7d30969d45fa0773883a37d730690f93&sort=plf-f&originationType=b&rr='
# snip web
snip5 = 'https://www.scopus.com/api/rest/sources/'

def get_main_words(page_source):
    mainWords0 = re.findall(r'<div class="sciTopicsVal displayNone"(.*?)</div>', page_source.text, re.S)
    mainWords = re.findall(r'"name":"(.*?)","id', str(mainWords0), re.S)
    return mainWords

def get_H(page_source):
    H = []
    hIndexUrls0 = re.findall(r'<section id="authorlist(.*?)</section>', page_source.text, re.S)
    hIndexUrls = re.findall(r'type="hidden"><a href="(.*?)" title="', str(hIndexUrls0), re.S)
    for hIndexUrl in hIndexUrls:
        hIndexUrl_source = requests.get(url=hIndexUrl, headers=headerDetail, allow_redirects=False)
        hIndex0 = re.findall(r'h</span>-index:(.*?)<button type=', hIndexUrl_source.text, re.S)
        hIndex = re.findall(r'<span class="fontLarge">(.*?)</span>', str(hIndex0), re.S)
        if hIndex:
            H.append(hIndex[0])
    return H

def get_SnipSjrRpNew(url):
    dataSnipSjrRp = []
    page_source = requests.get(url=url, headers=headers, allow_redirects=False)
    data0 = re.findall(r'<td data-type="source">\n<a href="(.*?)class="ddmDocSource"', page_source.text, re.S)
    if data0:
        data1 = data0[0]
        data2 = data1[10:22]
        snipUrl = snip5 + data2
        s1 = requests.Session()
        page_source1 = s1.get(snipUrl, headers=headers, allow_redirects = False)
        datasnip = re.findall(r'<name>SNIP</name><value>(.*?)</value>', page_source1.text, re.S)
        if datasnip:
            dataSnipSjrRp.append(datasnip[0])
            print(dataSnipSjrRp)
        else:
            dataSnipSjrRp.append('')

        datasjr = re.findall(r'<name>SJR</name><value>(.*?)</value>', page_source1.text, re.S)
        if datasjr:
            dataSnipSjrRp.append(datasjr[0])
        else:
            dataSnipSjrRp.append('')

        datarp = re.findall(r'<name>RP</name><value>(.*?)</value>', page_source1.text, re.S)
        if datarp:
            dataSnipSjrRp.append(datarp[0])
        else:
            dataSnipSjrRp.append('')

        return dataSnipSjrRp

    else:
        return []

def get_SnipSjrRp(url):
    dataSnipSjrRp = []
    page_source = requests.get(url=url, headers=headers, allow_redirects=False)
    data0 = re.findall(r'<td data-type="source">\n<a href="(.*?)class="ddmDocSource"', page_source.text, re.S)
    if data0:
        data1 = data0[0]
        data2 = data1[10:20]
        snipUrl = snip5 + data2
        s1 = requests.Session()
        page_source1 = s1.get(snipUrl, headers=headers, allow_redirects = False)
        print(page_source1.text)
        print(snipUrl)
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
        else:
            dataSnipSjrRp.append('')

        datarp0 = re.findall(r'name>RP&lt;/name>(.*?)/value>', page_source1.text, re.S)
        datarp = re.findall(r'&lt;value>(.*?)&lt;', str(datarp0), re.S)
        if datarp:
            dataSnipSjrRp.append(datarp[0])
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
    file = "../sourceData/a6-10.xls"
    data = xlrd.open_workbook(file, formatting_info=True)
    table = data.sheet_by_name('1')
    papers = []
    for i in range(76,3779):
        paper = {}
        content = table.row_values(i)
        paper['标题'] = content[4]
        paper['年份'] = content[5]
        paper['来源出版物名称'] = content[6]
        paper['DOI'] = content[14]
        paper['EID'] = content[44]
        print(paper['标题'])
        papers.append(paper)
    return papers

if __name__ == '__main__':

    papers = get_excel()
    fSub = open('../data/subject.txt', 'w')
    fCoun = open('../data/country.txt', 'w')
    fHIdex = open('../data/hIdex.txt', 'w')
    fCitation = open('../data/citation.txt', 'w')
    fMainWords = open('../data/mainWords.txt', 'w')
    fSnip = open('../data/snip.txt', 'w')
    fSjr = open('../data/sjr.txt', 'w')
    fRp = open('../data/rp.txt', 'w')
    fAffi = open('../data/affiliation.txt', 'w')

    for paper in papers:
        eid = paper['EID']
        doi = paper['DOI']

        # ====================================1=====================================

        if doi != '':
            url = url1 + doi + url2 + doi + url3 + doi + url4
            page_source = requests.get(url=url, headers=headers, allow_redirects=False)
            detailUrl1 = re.findall(r'<td data-type="docTitle">(.*?)</td>', page_source.text, re.S)
            detailUrl2 = re.findall(r'href="(.*?)"class="ddmDocTitle"', str(detailUrl1), re.S)

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
            if detailUrl2:
                detailUrl = detailUrl2[0].replace('amp;', '')
                page_source = requests.get(url=detailUrl, headers=headerDetail, allow_redirects=False)
                # -----------------H指数-------------------
                H = get_H(page_source)
                if H:
                    hIndexs = []
                    for h in H:
                        hindex = int(h)
                        hIndexs.append(hindex)
                    maxH = max(hIndexs)
                    print(maxH, file=fHIdex, flush=True)
                    print("H Index:")
                    print(maxH)
                    print("\n")
                else:
                    print('', file=fHIdex, flush=True)
                # -----------------主题词-------------------
                mainWords = get_main_words(page_source)
                print(mainWords, file=fMainWords, flush=True)
                print("Main Words:")
                print(mainWords)
                print("\n")
            else:
                print('', file=fHIdex, flush=True)
                print('', file=fMainWords, flush=True)
            time.sleep(1)


        else:
            print('', file=fHIdex, flush=True)
            print('', file=fMainWords, flush=True)
            print('', file=fSub, flush=True)
            print('', file=fCoun, flush=True)


        # ====================================2=====================================

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
                url='https://www.scopus.com/standard/retrieveClusterAttributes.uri', data=data,
                headers=headers)
            affiliation = re.findall(r'class="btnText">(.*?)</span>', rep.text, re.S)
            print("\nAffiliation:")
            print(affiliation)
            print(affiliation, file=fAffi, flush=True)
            time.sleep(1)
        else:
            print('', file=fAffi, flush=True)

        # ====================================3=====================================
        flag = 0
        if doi != '':
            # --------------SNIP SJR RP----------------
            doiSnip = doi.replace('/', '%2F')
            urlSnip0 = snip1 + doiSnip + snip2 + doiSnip + snip3 + doiSnip + snip4
            snipSjrRp = get_SnipSjrRpNew(urlSnip0)
            if snipSjrRp == [] and flag < 3:
                snipSjrRp = get_SnipSjrRpNew(urlSnip0)
                flag = flag + 1
            if snipSjrRp:
                snip = snipSjrRp[0]
                sjr = snipSjrRp[1]
                rp = snipSjrRp[2]
                print(snip, file=fSnip, flush=True)
                print(sjr, file=fSjr, flush=True)
                print(rp, file=fRp, flush=True)

                print("SNIP:" + snip + "  SJR:" + sjr + "  RJ:" + rp)
                print("\n")
            else:
                print('', file=fSnip, flush=True)
                print('', file=fSjr, flush=True)
                print('', file=fRp, flush=True)
            time.sleep(1)

        else:
            print('', file=fSnip, flush=True)
            print('', file=fSjr, flush=True)
            print('', file=fRp, flush=True)

        print("=======================================================================")
