# ----------------------正经的-----------------------
import xlrd
from selenium import webdriver
import time

url1 = 'https://wwww.scopus.com/results/results.uri?numberOfFields=0&src=s&clickedLink=&edit=&editSaveSearch=&origin=searchbasic&authorTab=&affiliationTab=&advancedTab=&scint=1&menu=search&tablin=&searchterm1='
url2 = '&field1=DOI&dateType=Publication_Date_Type&yearFrom=Before+1960&yearTo=Present&loadDate=7&documenttype=All&accessTypes=All&resetFormLink=&st1='
url3 = '&st2=&sot=b&sdt=b&sl=22&s=DOI%28'
url4 = '%29&sid=04d9016932b494f613c131f956db3e87&searchId=04d9016932b494f613c131f956db3e87&txGid=cd4965c8585c87e1b28188380fa2685e&sort=plf-f&originationType=b&rr='


def get_excel():
    file = "../sourceData/a种子文献.xls"
    data = xlrd.open_workbook(file, formatting_info=True)
    table = data.sheet_by_name('1')
    papers = []
    for i in range(11, 51):
        paper = {}
        content = table.row_values(i)
        paper['编号'] = content[0]
        paper['DOI'] = content[13]
        print(paper['编号'])
        print(paper['DOI'])
        papers.append(paper)
    return papers

if __name__ == '__main__':
    papers = get_excel()  # 获取表格中的数据

    for paper in papers:
        doi = paper['DOI']
        number = paper['编号']
        if doi != '':
            url = url1 + doi + url2 + doi + url3 + doi + url4
            print(number)
            print(url)
        else:
            print("The overflow！")
