import requests
import io
import xlrd
import time
import re
import sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='gb18030')

headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Cookie': 'UM_distinctid=16f03aeda951f7-0e445c640904db-12316b5a-13c680-16f03aeda96100; scopus.machineID=42A3D9C620FA39CD631B1EEC603CF6C6.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; __aza_perm=CheckPermissionCookie; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; xmlHttpRequest=true; everest_g_v2=g_surferid~XhQr8gAAFPW8EsFc; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; BIDUPSID=AD4E635AC5395DE343B6AC26FE46C873; PSTM=1580807141; BAIDUID=AD4E635AC5395DE32BF32FAB36D9518E:FG=1; _abck=DBA7C3BD927B70F835EE71C7C1AF1E5C~-1~YAAQ3tfSF1Al/c5vAQAALYYZEAP3PsVEykKgNX1Z7yztci5hDgw84xn4PLHgJApP9czi1665gl9UYf3Psj8hMewsC0VJyJ4zeYaAFnDRWBRn1V6QHFRN+AM1fdPLrdjVH03rIq8N/2uegBrLK1oqeHfP7xekO3huLcT5DQyMv7AsPC8sgBewMpaTC0r3cZ/sEsFtnKPxLxmLcpMThl0AM4o5v29/QwTJkACVqtbhSrQLE82F7JdGj0DDrhRbGu2lv036FerPg1M9d2o85I5lgfJvg6xNIT42/pvnc0U=~-1~-1~-1; sp=b75e6496-3e3c-4c58-acd1-2132ecf4d70c; _hjid=941e1bc6-293c-457e-bc91-4f5fc262dc09; _sp_id.9639=2314dd85-afcb-477b-9e82-64f1b98214d6.1580818111.2.1580906938.1580818818.be135e72-19df-41f8-acff-175cc2874545; s_vi=[CS]v1|2F1F44F18515E7A7-600006E4436E96C5[CE]; demdex=31931497131425969634038488604199850452; __cfduid=d09434e0603d6a09f49dbe43c5fe05fe11583840046; check=true; javaScript=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; screenInfo="900:1440"; SCSessionID=120CD8F4C3C824C75E991A6D23D9A5F6.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=effc8146-14bc-4deb-b; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB16789BD8FB731B14DD3D97925DD42DE028BADCB8A0E3BFEB304D004538394A4D610BA32070D9964CEACBAE7C5777723B783435B7E38EE30CCC67CAEF897963C26; _webvpn_key=eyJhbGciOiJIUzI1NiJ9.eyJ1c2VyIjoiMjAxNzIyNDQ5MiIsImlhdCI6MTU4NTgwMDY3NCwiZXhwIjoxNTg1ODg3MDc0fQ.kp-6hn3biepQqIi_z6GObRmNqdd9pVlCeSFTW6jlanI; webvpn_username=2017224492%7C1585800674%7Cdbdabe13a291db2b577efa7edd55c37ec9d587b1; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1266252805%7CMCIDTS%7C18355%7CMCMID%7C10178513813538045854600690115189903694%7CMCAAMLH-1586405524%7C11%7CMCAAMB-1586405524%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1585807925s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-18276%7CvVersion%7C4.4.1%7CMCCIDH%7C1246701492; s_pers=%20v8%3D1585800727664%7C1680408727664%3B%20v8_s%3DLess%2520than%25201%2520day%7C1585802527664%3B%20c19%3Dsc%253Asearch%253Adocument%2520searchform%7C1585802527712%3B%20v68%3D1585800702495%7C1585802527761%3B; s_sess=%20s_cpc%3D0%3B%20c21%3Ddoi%252810.1364%252Fol.30.001944%2529%3B%20e13%3Ddoi%252810.1364%252Fol.30.001944%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20e78%3Ddoi%252810.1364%252Fol.30.001944%2529%3B%20s_ppvl%3Dsc%25253Arecord%25253Asource%252520info%252C10%252C10%252C150%252C1424%252C150%252C1440%252C900%252C1%252CP%3B%20s_cc%3Dtrue%3B%20e41%3D1%3B%20s_sq%3Delsevier-sc-prod%25252Celsevier-global-prod%253D%252526c.%252526a.%252526activitymap.%252526page%25253Dsc%2525253Asearch%2525253Adocument%25252520searchform%252526link%25253DSearch%252526region%25253DsearchBtnRow%252526pageIDType%25253D1%252526.activitymap%252526.a%252526.c%252526pid%25253Dsc%2525253Asearch%2525253Adocument%25252520searchform%252526pidt%25253D1%252526oid%25253Dfunctiononclick%25252528event%25252529%2525257BsetCheckEmptyBasic%25252528true%25252529%2525253B%2525257D%252526oidt%25253D2%252526ot%25253DSUBMIT%3B%20s_ppv%3Dsc%25253Asearch%25253Adocument%252520searchform%252C41%252C31%252C560%252C1424%252C150%252C1440%252C900%252C1%252CP%3B; optimizelyPendingLogEvents=%5B%5D; optimizelyDomainTestCookie=0.9286843948461023; optimizelyDomainTestCookie=0.08768044853779289; optimizelyDomainTestCookie=0.8205166057476192; mbox=PC#28144c226fe547c38b225004f77d29fb.22_0#1649045532|session#4eee4e47c31044099a10a3cf86ad0e78#1585802487',
        'Host': 'www-scopus-com-s.webvpn.nefu.edu.cn',
        # 'Referer': 'https://www-scopus-com-s.webvpn.nefu.edu.cn/search/form.uri?display=basic&zone=header&origin=',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-site',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36',
}
headerDetail = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Cookie':'UM_distinctid=16f03aeda951f7-0e445c640904db-12316b5a-13c680-16f03aeda96100; scopus.machineID=42A3D9C620FA39CD631B1EEC603CF6C6.wsnAw8kcdt7IPYLO0V48gA; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelySegments=%7B%22278797888%22%3A%22gc%22%2C%22278846372%22%3A%22false%22%2C%22278899136%22%3A%22none%22%2C%22278903113%22%3A%22referral%22%7D; optimizelyBuckets=%7B%7D; __aza_perm=CheckPermissionCookie; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; xmlHttpRequest=true; everest_g_v2=g_surferid~XhQr8gAAFPW8EsFc; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; optimizelyEndUserId=oeu1578380260234r0.24122091685022462; BIDUPSID=AD4E635AC5395DE343B6AC26FE46C873; PSTM=1580807141; BAIDUID=AD4E635AC5395DE32BF32FAB36D9518E:FG=1; _abck=DBA7C3BD927B70F835EE71C7C1AF1E5C~-1~YAAQ3tfSF1Al/c5vAQAALYYZEAP3PsVEykKgNX1Z7yztci5hDgw84xn4PLHgJApP9czi1665gl9UYf3Psj8hMewsC0VJyJ4zeYaAFnDRWBRn1V6QHFRN+AM1fdPLrdjVH03rIq8N/2uegBrLK1oqeHfP7xekO3huLcT5DQyMv7AsPC8sgBewMpaTC0r3cZ/sEsFtnKPxLxmLcpMThl0AM4o5v29/QwTJkACVqtbhSrQLE82F7JdGj0DDrhRbGu2lv036FerPg1M9d2o85I5lgfJvg6xNIT42/pvnc0U=~-1~-1~-1; sp=b75e6496-3e3c-4c58-acd1-2132ecf4d70c; _hjid=941e1bc6-293c-457e-bc91-4f5fc262dc09; _sp_id.9639=2314dd85-afcb-477b-9e82-64f1b98214d6.1580818111.2.1580906938.1580818818.be135e72-19df-41f8-acff-175cc2874545; s_vi=[CS]v1|2F1F44F18515E7A7-600006E4436E96C5[CE]; demdex=31931497131425969634038488604199850452; __cfduid=d09434e0603d6a09f49dbe43c5fe05fe11583840046; check=true; javaScript=true; AMCVS_4D6368F454EC41940A4C98A6%40AdobeOrg=1; screenInfo="900:1440"; SCSessionID=120CD8F4C3C824C75E991A6D23D9A5F6.wsnAw8kcdt7IPYLO0V48gA; scopusSessionUUID=effc8146-14bc-4deb-b; AWSELB=CB9317D502BF07938DE10C841E762B7A33C19AADB16789BD8FB731B14DD3D97925DD42DE028BADCB8A0E3BFEB304D004538394A4D610BA32070D9964CEACBAE7C5777723B783435B7E38EE30CCC67CAEF897963C26; _webvpn_key=eyJhbGciOiJIUzI1NiJ9.eyJ1c2VyIjoiMjAxNzIyNDQ5MiIsImlhdCI6MTU4NTgwMDY3NCwiZXhwIjoxNTg1ODg3MDc0fQ.kp-6hn3biepQqIi_z6GObRmNqdd9pVlCeSFTW6jlanI; webvpn_username=2017224492%7C1585800674%7Cdbdabe13a291db2b577efa7edd55c37ec9d587b1; AMCV_4D6368F454EC41940A4C98A6%40AdobeOrg=1266252805%7CMCIDTS%7C18355%7CMCMID%7C10178513813538045854600690115189903694%7CMCAAMLH-1586405524%7C11%7CMCAAMB-1586405524%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1585807925s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-18276%7CvVersion%7C4.4.1%7CMCCIDH%7C1246701492; optimizelyDomainTestCookie=0.8591643540819034; optimizelyDomainTestCookie=0.16565117813516972; optimizelyDomainTestCookie=0.7771197815509847; mbox=PC#28144c226fe547c38b225004f77d29fb.22_0#1649045549|session#4eee4e47c31044099a10a3cf86ad0e78#1585802487; s_pers=%20v8%3D1585800754879%7C1680408754879%3B%20v8_s%3DLess%2520than%25201%2520day%7C1585802554879%3B%20c19%3Dsc%253Arecord%253Adocument%2520record%7C1585802555005%3B%20v68%3D1585800747201%7C1585802555215%3B; s_sess=%20s_cpc%3D0%3B%20c21%3Dtitle%2528color%2529%3B%20e13%3Dtitle%2528color%2529%253A1%3B%20c13%3Ddate%2520%2528newest%2529%3B%20e78%3Dtitle%2528color%2529%3B%20s_sq%3D%3B%20e41%3D1%3B%20s_cc%3Dtrue%3B%20s_ppvl%3Dsc%25253Asearch%25253Adocument%252520searchform%252C19%252C19%252C269%252C1424%252C232%252C1440%252C900%252C1%252CP%3B%20s_ppv%3Dsc%25253Arecord%25253Adocument%252520record%252C12%252C12%252C793%252C1424%252C451%252C1440%252C900%252C1%252CP%3B',
        'Host': 'www-scopus-com-s.webvpn.nefu.edu.cn',
        # 'Referer': 'https://www-scopus-com-s.webvpn.nefu.edu.cn/search/form.uri?display=basic&zone=header&origin=',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36',

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
detail4 = '&st2=&sid=bf43c8536b4c20bfb146383e8ae724aa&sot=b&sdt=b&sl=30&s=DOI%2810.'
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


# def get_2016_2017(s):
#     yin2016 = 0
#     yin2017 = 0
#     have2016 = s.find(name="li", attrs={"id": "li_2016"})
#
#     if have2016 != None:
#         yin2016 = have2016.find(name='span', attrs={'class': 'badge'}).find(name='span', attrs={'class': 'btnText'})
#         # print(yin2016)
#         yin2016 = locale.atoi(yin2016.text.replace(',',''))
#
#     have2017 = s.find(name="li", attrs={"id": "li_2017"})
#     if have2017 != None:
#         yin2017 = have2017.find(name='span', attrs={'class': 'badge'}).find(name='span', attrs={'class': 'btnText'})
#         yin2017 = locale.atoi(yin2017.text.replace(',',''))
#     return yin2016+yin2017

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
        hIndex = re.findall(r'<div><span class="fontLarge">(.*?)</span>', str(hIndex0), re.S)
        if hIndex:
            H.append(hIndex[0])
    return H

def get_SnipSjrRpNew(url):
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
    file = "../sourceData/1-5no.xls"
    data = xlrd.open_workbook(file, formatting_info=True)
    table = data.sheet_by_name('Sheet1')
    papers = []
    for i in range(344,725):
    # for i in range(262,264):
        paper = {}
        content = table.row_values(i)
        paper['标题'] = content[6]
        paper['年份'] = content[7]
        paper['来源出版物名称'] = content[8]
        paper['DOI'] = content[16]
        paper['EID'] = content[46]
        print(paper['标题'])
        print(paper['年份'])
        print(paper['来源出版物名称'])
        print(paper['DOI'])
        print(paper['EID'])
        papers.append(paper)

    # file = "../sourceData/1-10error.xls"
    # data = xlrd.open_workbook(file, formatting_info=True)
    # table = data.sheet_by_name('Sheet1')
    # papers = []
    # for i in range(1, 8):
    #     paper = {}
    #     content = table.row_values(i)
    #     paper['标题'] = content[4]
    #     paper['年份'] = content[5]
    #     paper['来源出版物名称'] = content[6]
    #     paper['DOI'] = content[8]
    #     paper['EID'] = content[13]
    #     print(paper['标题'])
    #     # print(paper['年份'])
    #     # print(paper['来源出版物名称'])
    #     # print(paper['DOI'])
    #     # print(paper['EID'])
    #     papers.append(paper)
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

        # # ====================================1=====================================

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
        #         detailUrl = detailUrl2[0].replace('amp;','')
        #         page_source = requests.get(url=detailUrl, headers=headerDetail, allow_redirects=False)
        #         # -----------------H指数-------------------
        #         H = get_H(page_source)
        #         if H:
        #             hIndexs = []
        #             for h in H:
        #                 hindex = int(h)
        #                 hIndexs.append(hindex)
        #             maxH = max(hIndexs)
        #             print(maxH, file=fHIdex, flush=True)
        #             print("H Index:")
        #             print(maxH)
        #             print("\n")
        #         else:
        #             print('', file=fHIdex, flush=True)
        #         # -----------------主题词-------------------
        #         mainWords = get_main_words(page_source)
        #         print(mainWords, file=fMainWords, flush=True)
        #         print("Main Words:")
        #         print(mainWords)
        #         print("\n")
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

        # # # ====================================2=====================================
        # #
        # # if eid != '':
        # #     url = citationturl1 + eid + citationturl2
        # #     s = serch(url, headers)
        # #
        # #     # --------------2016_2017引用次数---------------
        # #     citation = get_2016_2017(s)
        # #     print(citation, file=fCitation, flush=True)
        # #     print("2016_2017:")
        # #     print(citation)
        # # else:
        # #     print('', file=fCitation, flush=True)
        # # #
        #
        # ====================================3=====================================

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

        # ====================================4=====================================
        # if doi != '':
        #     # --------------SNIP SJR RP----------------
        #     doiSnip = doi.replace('/', '%2F')
        #     urlSnip0 = snip1 + doiSnip + snip2 + doiSnip + snip3 + doiSnip + snip4
        #     snipSjrRp = get_SnipSjrRpNew(urlSnip0)
        #     if snipSjrRp:
        #         snip = snipSjrRp[0]
        #         sjr = snipSjrRp[1]
        #         rp = snipSjrRp[2]
        #         print(snip, file=fSnip, flush=True)
        #         print(sjr, file=fSjr, flush=True)
        #         print(rp, file=fRp, flush=True)
        #
        #         print("SNIP:" + snip + "  SJR:" + sjr + "  RJ:" + rp)
        #         print("\n")
        #     else:
        #         print('', file=fSnip, flush=True)
        #         print('', file=fSjr, flush=True)
        #         print('', file=fRp, flush=True)
        #     time.sleep(1)
        #
        # else:
        #     print('', file=fSnip, flush=True)
        #     print('', file=fSjr, flush=True)
        #     print('', file=fRp, flush=True)

        print("=======================================================================")
