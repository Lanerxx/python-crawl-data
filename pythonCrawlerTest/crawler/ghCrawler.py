import urllib.request
from bs4 import BeautifulSoup
import xlrd
from selenium import webdriver
import time

ghURL1 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/record/display.uri?eid='
# 2-s2.0-84872975676
ghURL2 = '&origin=resultslist&sort=cp-t&src=s&st1='
# What+is+societal+impact+of+research+and+how+can+it+be+assessed+a+literature+survey
ghURL3 = '&st2=&sid=74f1a3ed29ef12039b0166fa3449211b&sot=b&sdt=b&sl=155&s=TITLE%28'
# What+is+societal+impact+of+research+and+how+can+it+be+assessed+a+literature+survey

def serch(url,headers):
    req = urllib.request.Request(url=url, headers=headers)
    rsp = urllib.request.urlopen(req,timeout=20000)
    html = rsp.read().decode()
    s = BeautifulSoup(html, 'html.parser')
    return s

def get_excel():
    file = "./s4.xls"

    data = xlrd.open_workbook(file, formatting_info=True)
    table = data.sheet_by_name('scopus')
    papers = []
    for i in range(13, 28):
        paper = {}
        content = table.row_values(i)
        paper['标题'] = content[2]
        paper['EID'] = content[4]
        print(paper['EID'])
        papers.append(paper)
    return papers

def get_url(eid, arTitle):
    arTitle = arTitle.replace('?', '')
    arTitle = arTitle.replace(' ', '+')
    arTitle = arTitle.replace('&', '%26')

    print(arTitle)
    url = ghURL1 + eid + ghURL2 + arTitle + ghURL3 + arTitle + '%'
    return url


if __name__ == '__main__':
    papers = get_excel()
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
        eid = paper['EID']
        arTitle = paper['标题']

        if eid != '':
            url = get_url(eid, arTitle)
            print(url)
            browser.get(url)
            if i == 1:

                time.sleep(1)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"referenceSrhResults\"]/span[1]").click()

                time.sleep(2)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"selectAllCheck\"]/label").click()
                time.sleep(1)

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"export_results\"]/span").click()

                time.sleep(1)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"exportList\"]/li[4]/label").click()

                time.sleep(1)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"citationGroupCheckboxes\"]/span/label").click()

                time.sleep(1)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"citationCheckBoxes\"]/ul/li[1]/label").click()

                time.sleep(1)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"authorIdChckbox\"]/label").click()

                time.sleep(1)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"citationCheckBoxes\"]/ul/li[3]/label").click()

                time.sleep(1)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"eidChckbox\"]/label").click()

                time.sleep(1)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"exportTrigger\"]/span").click()

                i = i + 1
            else:
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"referenceSrhResults\"]/span[1]").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"selectAllCheck\"]/label").click()

                time.sleep(1)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"directExport\"]/span").click()

        else:
            print("无EID")

    browser.close()