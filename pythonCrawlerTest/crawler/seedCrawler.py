# ----------------------正经的-----------------------
import xlrd
from selenium import webdriver
import time

url1 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/results/results.uri?numberOfFields=0&src=s&clickedLink=&edit=&editSaveSearch=&origin=searchbasic&authorTab=&affiliationTab=&advancedTab=&scint=1&menu=search&tablin=&searchterm1='
url2 = '&field1=DOI&dateType=Publication_Date_Type&yearFrom=Before+1960&yearTo=Present&loadDate=7&documenttype=All&accessTypes=All&resetFormLink=&st1='
url3 = '&st2=&sot=b&sdt=b&sl=22&s=DOI%28'
url4 = '%29&sid=04d9016932b494f613c131f956db3e87&searchId=04d9016932b494f613c131f956db3e87&txGid=cd4965c8585c87e1b28188380fa2685e&sort=plf-f&originationType=b&rr='

def myerror():
    time.sleep(1)
    try:
        browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
    except Exception:
        print('0')

def get_excel():
    file = "../sourceData/seed.xls"
    data = xlrd.open_workbook(file, formatting_info=True)
    table = data.sheet_by_name('1')
    papers = []
    for i in range(10, 11):
        paper = {}
        content = table.row_values(i)
        paper['DOI'] = content[12]
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
    time.sleep(1)
    i = 1
    for paper in papers:
        doi = paper['DOI']
        if doi != '':
            url = url1 + doi + url2 + doi + url3 + doi + url4
            print(url)
            browser.get(url)
            if i == 1:
                myerror()
                browser.find_element_by_xpath("//*[@id=\"selectAllCheck\"]/label").click()
                myerror()
                browser.find_element_by_xpath("//*[@id=\"moreOptionToggleBtn\"]/span[2]").click()
                myerror()
                browser.find_element_by_xpath("//*[@id=\"moreOptionsMenu\"]/div/span[2]/span/ul/li[1]/button/span").click()
                time.sleep(10)
                myerror()
                browser.find_element_by_xpath("//*[@id=\"selectAllCheck\"]/label").click()
                myerror()
                browser.find_element_by_xpath("//*[@id=\"export_results\"]/span").click()
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
                time.sleep(3)
                i = i + 1
            else:
                myerror()
                browser.find_element_by_xpath("//*[@id=\"moreOptionToggleBtn\"]/span[2]").click()
                myerror()
                browser.find_element_by_xpath("//*[@id=\"moreOptionsMenu\"]/div/span[2]/span/ul/li[1]/button/span").click()
                time.sleep(10)
                myerror()
                browser.find_element_by_xpath("//*[@id=\"selectAllCheck\"]/label").click()
                myerror()
                browser.find_element_by_xpath("//*[@id=\"directExport\"]/span").click()
                time.sleep(3)
            print("---------------------")
        else:
            print("The overflow！")
    time.sleep(2)
    browser.close()