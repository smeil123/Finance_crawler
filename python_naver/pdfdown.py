import sys
from pandas import Series, DataFrame
from openpyxl import Workbook
import pandas as pd
import os
from urllib import request

## 프로그래스바
def progressBar(value, endvalue, bar_length=20):
    percent = float(value) / endvalue
    arrow = '-' * int(round(percent * bar_length)-1) + '>'
    spaces = ' ' * (bar_length - len(arrow))

    sys.stdout.write("\rPercent: [{0}] {1}%".format(arrow + spaces, int(round(percent * 100))))
    sys.stdout.flush()

## pdf다운로드 함수
# - data : excel데이터
# - dir_path : pdf를 저장할 경로
def pdf_download(data,dir_path=None):

    pdf_list = []

    ## pdf 경로를 지정해주지 않으면 임의의 경로에 저장
    if dir_path is None:
        down_path = "./pdf"
    else:
        down_path = dir_path

    ## 다운받을 경로폴더 생성
    try:
        os.mkdir(down_path)
    except:
        print("폴더 있음")

    for i,pdf_url in enumerate(data["pdf"]):
        progressBar(i,len(data["pdf"]))
        try:
            filename = os.path.join(down_path,pdf_url.split('/')[-1])
        except:
            #pdf가 없는 경우
            pass
        try:
            ## url로 이동해서 pdf다운
            request.urlretrieve(pdf_url,filename)
        except:
            ## url이 잘못된 경우
            pdf_list.append(pdf_url)
            print("예외 발생")
    return pdf_list

if __name__ == "__main__":
    excelfile = input("excel파일명 : ")
    downpath = input("pdf 저장받을 위치( n - ./pdf): ")

    if downpath == 'n':
        downpath = None

    ## excel파일 열기
    excel_data = pd.read_excel(excelfile)
    ## pdf다운
    pdfs = pdf_download(excel_data,downpath)

    ## 잘못된 url이 있는 목록 저장
    if pdfs is not None:
        with io.open(os.path.join("다운받지못한pdf"+".txt"),'w',encoding='utf8') as f:
            for pt in pdfs:
                f.write(pt)
                f.write("\n")
            f.close()
