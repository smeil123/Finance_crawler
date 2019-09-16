import pandas as pd
from pandas import Series, DataFrame
import io
import os
import re

## pdf를 읽어서 txt파일로 변환한 파일들을 엑셀파일로 합치는 코드

def read_text(dirname):
    # 디렉토리에 있는 파일이름 읽음
    text_list = os.listdir(dirname)
    pdf_txt = []

    for text in text_list:
        # txt파일만
        if '.txt' == os.path.splitext(text)[-1]:
            with io.open(os.path.join(dirname,text),'r',encoding='utf8') as f:
                text_data = ''
                # txt파일 read
                while(1):
                    line = f.readline()
                    if line:
                        # 단어만 존재 or 공백만 존재하면 제외
                        if len(re.findall("[\S]",line)) == 0:
                            pass
                        if len(re.findall("[\ ]",line)) < 2:
                            pass
                        else:
                            text_data = text_data + line
                    else:
                        break

            pdfname = os.path.splitext(text)[0]
            pdf_txt.append([pdfname,text_data])

    # 데이터프레임으로 변환
    pdf_txt = DataFrame(pdf_txt)
    pdf_txt.columns = ["pdf명","text"]

    return pdf_txt

if __name__ == "__main__":
    dirname = input("txt폴더 : ")
    excel_path = input("데이터를 합칠 excel파일 :")

    process_textdata = read_text(dirname)

    excel_data = pd.read_excel(excel_path)
    ## 추기된 부분
    # pdfurl이 없는 경우 nan(not a number)로 float형으로 저장되기 때문에 str변환이 필요
    excel_data['pdf명'] = [str(i).split('/')[-1] for i in excel_data['pdf']]
    excel_data

    data = pd.merge(excel_data,process_textdata,how = 'left')
    writer = pd.ExcelWriter(os.path.splitext(excel_path)[0]+"_text.xlsx")
    data.to_excel(writer,'Sheet1')
    writer.save()
