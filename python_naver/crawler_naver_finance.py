# 네이버금융 > 종목분석 url : https://finance.naver.com/research/company_list.nhn
## chromedriver.exe 경로 필요 :  http://chromedriver.chromium.org/downloads
## selenium을 사용해서 동적으로 웹크롤링을 진행한다
## selenium을 사용하면 원하는 정보의 xpath를 알아야한다 (크롬에서 f12로 확인)

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from pandas import Series, DataFrame
from openpyxl import Workbook
import pandas as pd
import os
import time
from urllib import request

# - driver_path : chromedriver.exe가 설치되어 있는 경로
# - path : 엑셀파일이 저장될 위치(None : 현재위치)
# - start_date, end_date : 검색기간 설정(None : 전체기간)
# - pageNum : 가져올 페이지 수 (None : 검색된 전체 페이지)

def naver_finace_crawling(driver_path,path = None,start_date=None,end_date=None,pageNum=None):
    try:
        driver = webdriver.Chrome(driver_path)
    except:
        print("chrome driver 경로를 확인해주세요")
        return False,False

    if path is None:
        file_path = "./"
    else:
        file_path = path
    try:
        os.mkdir(file_path)
        print("%s폴더 생성"%path)
    except:
        print("path 있음")


    driver.implicitly_wait(3)

    page = 1
    stock_list = [] # 데이터 저장
    url_date = "searchType=writeDate"
    url = 'https://finance.naver.com/research/company_list.nhn?'
    #end_flag = False;

    if start_date is not None:
        url = url + url_date + '&writeFromDate=' +start_date +"&writeToDate=" + end_date
        if end_date is None:
            print("start date가 있으면 end date도 필요합니다")
            return False, False

    while True:

        # 페이지수를 입력받은 경우
        if pageNum is not None:
            if pageNum == 0:
                break
            pageNum -= 1

        c_url = url +'&page=' + str(page)
        print(c_url)
        driver.get(c_url) # 해당 url로 이동

        html = driver.page_source

        # 테이블 정보 get
        table = driver.find_elements_by_xpath('//*[@id="contentarea_left"]/div[3]/table[1]/tbody/tr')

        for i,t in enumerate(table):

            # 테이블의 행을 읽는다. 이때 줄바꿈 행도 읽기 때문에 rows의 길이 확인
            rows = t.find_elements_by_tag_name('td')
            if len(rows) == 6:
                temp = ['0' for _ in range(6)] # 저장할 공간을 미리 할당

                for k,row in enumerate(rows):
                    temp[k] = row.text

                    # pdf경로는 tag가 다름
                    if row.get_attribute("class") == "file":
                        try:
                            temp[k] = row.find_element_by_tag_name("a").get_attribute('href')
                        except:
                            temp[k] = None

                stock_list.append(temp)

        # 마지막 페이지인 경우 stop
        pagination = driver.find_element_by_xpath('//*[@id="contentarea_left"]/div[3]/table[2]/tbody/tr/td[@class="on"]')
        if pagination.text != str(page):
            break
        page += 1

    driver.quit()
    # DataFrame으로 변환 -> excel저장 용이
    try :
        stocks = DataFrame(stock_list)
    except:
        print(stock_list)
        print("저장할 데이터가 없거나 형식이 잘못되었습니다")
        return False, stock_list

    stocks.columns = ["종목명","제목","증권사","pdf","작성일","조회수"]
    #엑셀로 저장
    # 시간정보를 파일명에 더해서 덮여쓰여지는 것을 방지(존재하는 파일명을 사용하면 덮어쓰기된다)
    file_path = os.path.join(file_path,"naver_fiance" +format(time.time(),'.0f')+ ".xlsx")
    print(file_path)

    try:
        writer = pd.ExcelWriter(file_path)
        stocks.to_excel(writer,'Sheet1')
        writer.save()
        print("%s에 저장했습니다" % file_path)

    except:
        print("\n저장하지 못했습니다")

    return file_path, stocks

if __name__ == "__main__":
    chromedriver = input("Chromedriver 위치 : ")
    excel_path = input("네이버 경제 데이터 저장받을 위치( n - 현재 위치): ")
    start_date = input("검색조건 - 시작날짜(yyyy-mm-dd) ( n  - all): ")
    end_date = input("검색조건 - 종료날짜(yyyy-mm-dd) ( n - all): ")
    pagenum = input("페이지 수 ( n - all) : ")

    if excel_path == 'n':
        excel_path = None

    if start_date == 'n':
        start_date = None

    if end_date == 'n':
        end_date = None

    if pagenum == 'n':
        pagenum = None
    else:
        pagenum = int(pagenum)

    filename, data = naver_finace_crawling(chromedriver,excel_path,start_date,end_date,pagenum)
