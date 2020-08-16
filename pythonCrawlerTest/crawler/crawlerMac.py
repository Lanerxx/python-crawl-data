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
        'cookie': 'scopus.machineID=565E426BB5379DC87C63C542901909C3.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1587562861883r0.38711136738597673; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; xmlHttpRequest=true; __cfduid=d2f7cd7f9481c1bd55976b0eb1883d03d1593781793; check=true; javaScript=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; screenInfo="900:1440"; __cfruid=c308da74070671083fc711f4995ffe799cc41397-1594198091; SCSessionID=F0FADF62A2CEF7FC1AE71850F4B09FD7.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=1860802a-7a13-4ed8-9; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB1DEE74CB13F0DD5AAF32FCA68545060DF04647BA62027A2595E6B5CC63FDEFD4FA31AAC5A6BDE3E4B4DACF34F3854CEEBCEB91DB25896B1DDEB7DEF7A530F30CD; optimizelyPendingLogEvents=%5B%5D; mbox=PC#7f4649f1b841468a940386570b585808.38_0#1657457790|session#73377b58181a4649a84718193c2d38dd#1594214820; s_pers=%20v8%3D1594212989942%7C1688820989942%3B%20v8_s%3DLess%2520than%25201%2520day%7C1594214789942%3B%20c19%3Dsc%253Asearch%253Adocument%2520searchform%7C1594214789959%3B%20v68%3D1594212987971%7C1594214789991%3B; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1075005958%7CMCIDTS%7C18452%7CMCMID%7C09635205066540933810671114848219444600%7CMCAAMLH-1594817790%7C11%7CMCAAMB-1594817790%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1594220190s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-18382%7CMCCIDH%7C1249006796%7CvVersion%7C4.4.1; s_sess=%20s_cpc%3D0%3B%20c21%3Dtitle-abs-key%2528color%2529%3B%20e13%3Dtitle-abs-key%2528color%2529%253A1%3B%20c13%3Dcited%2520by%2520%2528highest%2529%3B%20e78%3Dtitle-abs-key%2528color%2529%3B%20s_sq%3D%3B%20s_ppvl%3Dsc%25253Arecord%25253Adocument%252520record%252C13%252C13%252C315%252C1440%252C190%252C1440%252C900%252C1%252CP%3B%20e41%3D1%3B%20s_cc%3Dtrue%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520searchform%252C74%252C74%252C789%252C1440%252C190%252C1440%252C900%252C1%252CP%3B',
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
        'cookie': 'scopus.machineID=565E426BB5379DC87C63C542901909C3.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1587562861883r0.38711136738597673; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; xmlHttpRequest=true; __cfduid=d2f7cd7f9481c1bd55976b0eb1883d03d1593781793; SCSessionID=FB86B255F54D604EAAED50644CAC29A5.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=2a3b12b2-c8fd-441e-9; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB1EEB6D65004F87E82B9B085DF4C3E5607FDDD6CA1D2EDEFF38F58F68A4277CB2CBAFDF2ADE925350150D7900CAD0CA8A6CEB91DB25896B1DDEB7DEF7A530F30CD; check=true; javaScript=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1075005958%7CMCIDTS%7C18452%7CMCMID%7C09635205066540933810671114848219444600%7CMCAAMLH-1594802674%7C11%7CMCAAMB-1594802674%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1594205074s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-18382%7CMCCIDH%7C1249006796%7CvVersion%7C4.4.1; screenInfo="900:1440"; optimizelyPendingLogEvents=%5B%5D; __cfruid=c308da74070671083fc711f4995ffe799cc41397-1594198091; mbox=PC#7f4649f1b841468a940386570b585808.38_0#1657442902|session#0be1b713aeb04fecbacabdde913b19fa#1594199734; s_pers=%20v8%3D1594198104976%7C1688806104976%3B%20v8_s%3DLess%2520than%25201%2520day%7C1594199904976%3B%20c19%3Dsc%253Arecord%253Adocument%2520record%7C1594199904993%3B%20v68%3D1594198101312%7C1594199905030%3B; s_sess=%20s_cpc%3D0%3B%20c21%3Dtitle-abs-key%2528color%2529%3B%20e13%3Dtitle-abs-key%2528color%2529%253A1%3B%20c13%3Dcited%2520by%2520%2528highest%2529%3B%20e78%3Dtitle-abs-key%2528color%2529%3B%20s_sq%3D%3B%20e41%3D1%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Arecord%25253Adocument%252520record%252C19%252C19%252C471%252C1440%252C346%252C1440%252C900%252C1%252CP%3B%20s_ppv%3Dsc%25253Arecord%25253Adocument%252520record%252C19%252C19%252C471%252C1440%252C346%252C1440%252C900%252C1%252CP%3B',
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

def Cheak_main_words(mainWords,mainWordsChe):
    if mainWords ==mainWordsChe:
        return 1
    else:
        return 0

def get_main_words(page_source):
    mainWords0 = re.findall(r'<div class="sciTopicsVal displayNone"(.*?)</div>', page_source.text, re.S)
    mainWords = re.findall(r'"name":"(.*?)","id', str(mainWords0), re.S)
    return mainWords

def Cheak_H(HChe0,HChe1,HChe2):
    H = ['X']
    if HChe0 == HChe1 or HChe0 ==HChe2:
        H = HChe0
    if HChe1 ==HChe2:
        H = HChe1
    return H

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

def Cheak_SnipSjrRpNew(snipSjrRp,snipSjrRpChe):
    if snipSjrRp ==snipSjrRpChe:
        return 1
    else:
        return 0

def get_SnipSjrRpNew(url):
    dataSnipSjrRp = []
    page_source = requests.get(url=url, headers=headers, allow_redirects=False)
    data0 = re.findall(r'<td data-type="source">\n<a href="(.*?)class="ddmDocSource"', page_source.text, re.S)
    if data0:
        data1 = data0[0]
        data2 = data1[10:21]
        snipUrl = snip5 + data2
        s1 = requests.Session()
        page_source1 = s1.get(snipUrl, headers=headers, allow_redirects = False)
        datasnip = re.findall(r'<name>SNIP</name><value>(.*?)</value>', page_source1.text, re.S)
        if datasnip:
            dataSnipSjrRp.append(datasnip[0])
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
    print(data0)
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
    file = "../sourceData/1-25.xls"
    data = xlrd.open_workbook(file, formatting_info=True)
    table = data.sheet_by_name('1')
    papers = []
    for i in range(792,995):
        paper = {}
        content = table.row_values(i)
        paper['标题'] = content[4]
        paper['年份'] = content[5]
        paper['来源出版物名称'] = content[6]
        paper['DOI'] = content[8]
        paper['EID'] = content[13]
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

        # if doi != '':
        #     url = url1 + doi + url2 + doi + url3 + doi + url4
        #     page_source = requests.get(url=url, headers=headers, allow_redirects=False)
        #     detailUrl1 = re.findall(r'<td data-type="docTitle">(.*?)</td>', page_source.text, re.S)
        #     detailUrl2 = re.findall(r'href="(.*?)"class="ddmDocTitle"', str(detailUrl1), re.S)
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
        #     if detailUrl2:
        #         detailUrl = detailUrl2[0].replace('amp;', '')
        #         page_source = requests.get(url=detailUrl, headers=headerDetail, allow_redirects=False)
        #         # -----------------H指数-------------------
        #         print("H Index:")
        #         HChe0 = get_H(page_source)
        #         time.sleep(1)
        #         HChe1 = get_H(page_source)
        #         time.sleep(1)
        #         HChe2 = get_H(page_source)
        #         H = Cheak_H(HChe0,HChe1,HChe2)
        #         if H != ['X']:
        #             if H:
        #                 hIndexs = []
        #                 for h in H:
        #                     hindex = int(h)
        #                     hIndexs.append(hindex)
        #                 maxH = max(hIndexs)
        #                 print(maxH, file=fHIdex, flush=True)
        #                 print(maxH)
        #                 print("\n")
        #             else:
        #                 print('',file=fHIdex,flush=True)
        #         else:
        #             print('ERROR', file=fHIdex, flush=True)
        #         # -----------------主题词-------------------
        #         print("Main Words:")
        #         mainWords = get_main_words(page_source)
        #         mainWordsCHe = get_main_words(page_source)
        #         flag = Cheak_main_words(mainWords,mainWordsCHe)
        #         if(flag ==1):
        #             print(mainWords,file=fMainWords,flush=True)
        #             print(mainWords)
        #             print("\n")
        #         else:
        #             print("ERROR",file=fMainWords,flush=True)
        #
        #     else:
        #         print('', file=fHIdex, flush=True)
        #         print('', file=fMainWords, flush=True)
        #     time.sleep(1)
        #
        #
        # else:
        #     print('', file=fHIdex, flush=True)
        #     print('', file=fMainWords, flush=True)
        #     print('', file=fSub, flush=True)
        #     print('', file=fCoun, flush=True)


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
        # if doi != '':
        #     # --------------SNIP SJR RP----------------
        #     doiSnip = doi.replace('/', '%2F')
        #     urlSnip0 = snip1 + doiSnip + snip2 + doiSnip + snip3 + doiSnip + snip4
        #     snipSjrRp = get_SnipSjrRpNew(urlSnip0)
        #     time.sleep(1)
        #     snipSjrRpChe = get_SnipSjrRpNew(urlSnip0)
        #     flag = Cheak_SnipSjrRpNew(snipSjrRp, snipSjrRpChe)
        #     if flag == 1:
        #         if snipSjrRp:
        #             snip = snipSjrRp[0]
        #             sjr = snipSjrRp[1]
        #             rp = snipSjrRp[2]
        #             print(snip, file=fSnip, flush=True)
        #             print(sjr, file=fSjr, flush=True)
        #             print(rp, file=fRp, flush=True)
        #
        #             print("SNIP:" + snip + "  SJR:" + sjr + "  RJ:" + rp)
        #             print("\n")
        #         else:
        #             print('', file=fSnip,flush=True)
        #             print('', file=fSjr,flush=True)
        #             print('', file=fRp,flush=True)
        #     else:
        #         print('ERROR', file=fSnip, flush=True)
        #         print('ERROR', file=fSjr, flush=True)
        #         print('ERROR', file=fRp, flush=True)
        #     time.sleep(1)
        #
        # else:
        #     print('', file=fSnip, flush=True)
        #     print('', file=fSjr, flush=True)
        #     print('', file=fRp, flush=True)

        print("=======================================================================")
