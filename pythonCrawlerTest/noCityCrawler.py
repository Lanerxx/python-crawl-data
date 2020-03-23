# ----------------------正经的-----------------------
import xlrd
from selenium import webdriver
import time

url1 = 'https://www-scopus-com-s.webvpn.nefu.edu.cn/results/results.uri?numberOfFields=0&src=s&clickedLink=&edit=&editSaveSearch=&origin=searchbasic&authorTab=&affiliationTab=&advancedTab=&scint=1&menu=search&tablin=&searchterm1='
# American+Journal+of+Preventive+Medicine
url2 = '&field1=SRCTITLE&dateType=Publication_Date_Type&yearFrom='
# 2008
url3 = '&yearTo='
# 2008
url4 = '&loadDate=7&documenttype=All&accessTypes=All&resetFormLink=&st1='
# American+Journal+of+Preventive+Medicine
url5 = '&st2=&sot=b&sdt=b&sl=25&s=SRCTITLE%28'
# American+Journal+of+Preventive+Medicine
url6 = '%29&sid=3538c4dc0bbee987eefd0909d22f53ce&searchId=3538c4dc0bbee987eefd0909d22f53ce&txGid=d5aba7ecda2b6b93905146f9fed5607d&sort=cp-t&originationType=b&rr='

def get_excel():
    file = "./1-25.xls"

    data = xlrd.open_workbook(file, formatting_info=True)
    table = data.sheet_by_name('引用文献')
    papers = []
    for i in range(3, 7):
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
        print(paper['DOI'])
        papers.append(paper)
    return papers

def get_url(year, sourceTitle):
    year = str(year)[0:4]
    print(year)
    sourceTitle = sourceTitle.replace(' ', '+')
    print(sourceTitle)
    url = url1 + sourceTitle + url2 + year + url3 + year + url4 + sourceTitle + url5 + sourceTitle + url6
    return url

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
        sourceTitle = paper['来源出版物名称']
        if year >= 1960:
            url = get_url(year, sourceTitle)
            browser.get(url)
            if i == 1:
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"resultDataRow0\"]/th/div/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"resultDataRow1\"]/th/div/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"resultDataRow2\"]/th/div/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"resultDataRow3\"]/th/div/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"chunkExport\"]").click()

                time.sleep(1)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"exportList\"]/li[4]/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"bibliographicalInformationCheckboxes\"]/span/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"abstractInformationCheckboxes\"]/span/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"fundInformationCheckboxes\"]/span/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"otherInformationCheckboxes\"]/span/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"otherInfoCheckboxes\"]/ul/li[4]/label").click()

                time.sleep(1)
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"exportTrigger\"]").click()
                i = i + 1
            else:
                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"resultDataRow0\"]/th/div/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"resultDataRow1\"]/th/div/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"resultDataRow2\"]/th/div/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"resultDataRow3\"]/th/div/label").click()

                try:
                    browser.find_element_by_xpath("//*[@id=\"_pendo-close-guide_\"]").click()
                except Exception:
                    print('0')
                browser.find_element_by_xpath("//*[@id=\"directExport\"]/span").click()



            print("---------------------")


        else:
            print("The overflow！")



    browser.close()
