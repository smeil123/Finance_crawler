import pandas as pd
from pandas import Series, DataFrame
import io
import os
from nltk.tokenize import sent_tokenize
from nltk.tokenize import sent_tokenize
import re

# analyst 를 찾기위해 전부 읽어들이기
# 이때는 전처리를 하지않고 읽기
def analyst_read_text(dirname):
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
                        text_data = text_data + line
                    else:
                        break

            pdfname = os.path.splitext(text)[0]
            pdf_txt.append([pdfname,text_data])

    # 데이터프레임으로 변환
    pdf_txt = DataFrame(pdf_txt)
    pdf_txt.columns = ["pdf명","text"]

    return pdf_txt

#analyst라고 명시되어 있는 분삭가의 찾기
def analyst_find(textdata):
    analysts = [None for _ in range(len(textdata))]
    analysts_check= [None for _ in range(len(textdata))]
    # text파일에서 이름이 포함된 부분을 추출
    for i,pdf in enumerate(textdata["text"]):

        # Analyst와 analyst 같이 대 소문자가 다를 수 있기 때문에
        pdf = str(pdf).lower() #소문자변환

        pdf = sent_tokenize(pdf) #문장단위 토큰화

        name = ''
        for sent in pdf:
            # analyst라는 단어를 찾으면 그 문장을 저장
            if sent.find("analyst") >=0:
                name = sent
                break

        #name을 못찾은 경우
        if name is None:
            print("없음")
            continue
        else:
            name = name.split('\n')

        L = []
        for j,n in enumerate(name):
            # analyst가 있는 주변을 추출
            if n.find('analyst') >=0:
                if len(name) == j+1:
                    L = [n,'']
                else:
                    L = [n,name[j+1]]

        analysts[i] = L

    # 이름이 아닌 부분을 제거
    pre_analysts = [None for _ in range(len(textdata))]

    for i, analyst in enumerate(analysts):
        if analyst is None:
            continue
        # 한글빼고 제거
        analyst = [re.sub("[^가-힣]",'',a) for a in analyst]

        for a in analyst:
            if a is not "":
                pre_analysts[i] =  a

                if len(pre_analysts[i]) == 6:
                    #사람이름+사람이름
                    analysts_check[i] = 'check'
                    pre_analysts[i] =  pre_analysts[i][:3]+','+pre_analysts[i][-3:]
                    continue

                if len(pre_analysts[i]) > 6:
                    # 회사명+사람이름
                    analysts_check[i] = 'check'
                    pre_analysts[i] = pre_analysts[i][-3:]

    # 데이터프레임 생성
    pre_analysts = DataFrame(pre_analysts)
    pre_analysts.columns = ["분석가"]

    analysts_check = DataFrame(analysts_check)
    analysts_check.columns = ["확인해 볼 분석가"]


    return pd.merge(pre_analysts,analysts_check,left_index=True,right_index=True)

## email형식을 찾아서 분석가 찾기
def analyst_email_find(textdata,analysts_df):

    # 정규표현식으로 이메일찾기
    p = re.compile('\w+[@]\w+[.]')

    for i,text in enumerate(textdata['text']):
        if analysts_df['분석가'][i] is None:
            a_email = p.search(str(text))
            if a_email is not None:
                # 이메일이 적힌 위치 앞
                name = text[a_email.start()-30:a_email.start()]
                analysts_df['분석가'][i] = re.sub("[^가-힣]",'',name)

                if len(analysts_df['분석가'][i]) > 3:
                    # 3글자 이상이면 회사명+사람이름 이라 생각하고 전처리
                    analysts_df["확인해 볼 분석가"][i] = 'check'
                    analysts_df['분석가'][i] = analysts_df['분석가'][i][-3:]

    return analysts_df


if __name__ == "__main__":
    dirname = input("txt폴더 : ")
    excel_path = input("데이터를 합칠 excel파일 :")

    text_data = analyst_read_text(dirname)

    df_analysts = analyst_find(text_data)
    df_analysts = analyst_email_find(text_data,df_analysts)

    # 데이터프레임을 하나로 합친다
    df_analyst_pdf = pd.merge(text_data,df_analysts,left_index=True,right_index=True)
    # 전처리 되지 않은 text는 저장할 필요가 없어서 제거
    del df_analyst_pdf['text']

    excel = pd.read_excel(excel_path)
    # pdf명을 기준으로 merge
    data = pd.merge(excel,df_analyst_pdf)
    # 엑셀저장
    writer = pd.ExcelWriter(os.path.splitext(excel_path)[0]+"_analyst.xlsx")
    data.to_excel(writer,'Sheet1')
    writer.save()
