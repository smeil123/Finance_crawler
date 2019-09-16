import pandas as pd
from pandas import Series, DataFrame
import sys
import datetime
import re
import nltk
import json
import os
from nltk.tokenize import sent_tokenize
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from konlpy.tag import Twitter

tagger = Twitter() # Twitter 태깅 함수

##  프로그레스바
import sys
def progressBar(value, endvalue, bar_length=20):
    percent = float(value) / endvalue
    arrow = '-' * int(round(percent * bar_length)-1) + '>'
    spaces = ' ' * (bar_length - len(arrow))

    sys.stdout.write("\rPercent: [{0}] {1}%".format(arrow + spaces, int(round(percent * 100))))
    sys.stdout.flush()

## 단어 수 세기
# - textdata : excel에 있는 text데이터
# - dir_path : pdf를 저장할 경로
def count_word(textdata):
    print("----------------------")
    print("단어 수 세는 중")
    process_text = [None for _ in range(len(textdata))]

    for idx in range(len(textdata)):
        # 읽지 못한 pdf파일은 건너뛴다
        if type(textdata[idx]) == str:
            sent_text = textdata[idx].lower()
            sent_text = sent_tokenize(sent_text)
            sent_text = [re.sub("\n",'',text) for text in sent_text]

            # 숫자 및 특수문자 제거
            #sent_text = [re.sub("[^(가-힣ㄱ-ㅎㅏ-ㅣa-zA-Z\ )]",' ',text) for text in sent_text]
            # 한글빼고 다 제거
            sent_text = [re.sub("[^(가-힣)]",' ',text) for text in sent_text]
            sent_text = [re.sub("[(]",' ',text) for text in sent_text]
            sent_text = [re.sub("[)]",' ',text) for text in sent_text]
            sent_text = [re.sub("\ +",' ',text) for text in sent_text]
            # 공백이 2개 이상일때
            sent_text = [re.sub("  +",' ',text) for text in sent_text]

            process_text[idx] = sent_text

    word_cnt = []
    for idx in range(len(process_text)):
        tmp = process_text[idx]
        if tmp is not None:
            text_token = [word_tokenize(text) for text in tmp]
            text_token = [y for x in text_token for y in x]
            word_cnt.append(len(text_token))
        else:
            word_cnt.append(0)

    return process_text,word_cnt
    ## word_cnt는 list -> return 받으면 dataframe으로 저장해야된다

## 긍부정 단어 수
# - textdata : 전처리한 text데이터
# - pn_dict : 긍부정 사전
def count_pos_neg(textdata,pn_dict=None):

    # 긍부정 사전 open
    if pn_dict is None:
        pn_dict = './dict.json'

    json1_file = open(pn_dict)
    json1_str = json1_file.read()
    pos_neg_dict = json.loads(json1_str)

    # text데이터에서 명사 추출(오래걸림)
    start = datetime.datetime.now()
    text_tag = []

    for sent_text in process_text:
        text_token = []

        if sent_text is None:
            text_tag.append(text_token)
            continue
        # 문장단위 토큰화에서 명사 추출
        for text in sent_text:
            # 여기서 비교의 용이를 위해 2음절 이상의 단어만 가져옴
            if len(text) >= 2:
                text_token += tagger.nouns(text)

        text_tag.append(text_token)
        progressBar(len(text_tag),len(process_text))

    end = datetime.datetime.now()
    print("\n걸린시간 : ",end-start)

    # 긍부정 단어 세기
    pos_cnt = []
    neg_cnt = []

    for text in text_tag:
        pos,neg = 0,0
        for word in text:
            if word in pos_neg_dict:
                np = pos_neg_dict.get(word)
                if np == "pos":
                    pos += 1
                elif np == "neg":
                    neg += 1
        pos_cnt.append(pos)
        neg_cnt.append(neg)

    return pos_cnt, neg_cnt

if __name__ == "__main__":
    excelfile = input("excel파일명 : ")
    data_flag = input("dart? naver? (d/n) :")
    pn_dict = input("긍부정사전( n - ./dict.json): ")
    ## excel파일 열기
    excel_data = pd.read_excel(excelfile)

    if pn_dict == 'n':
        pn_dict = None

    if data_flag == "d":
        process_text,word_cnt = count_word(excel_data["텍스트"])
    else:
        process_text,word_cnt = count_word(excel_data["text"])

    pos,neg = count_pos_neg(process_text,pn_dict)

    # 데이터 합치기
    df = DataFrame(word_cnt)
    pos_df = DataFrame(pos)
    neg_df = DataFrame(neg)

    data2 = pd.concat([df, pos_df, neg_df], axis=1)
    data2.columns = ["단어수","긍정단어수","부정단어수"]

    data = pd.merge(excel_data,data2,right_index=True,left_index=True)

    writer = pd.ExcelWriter(os.path.splitext(excelfile)[0]+'_wordcnt.xlsx')
    data.to_excel(writer,'Sheet1')
    writer.save()

    print("파일이 저장되었습니다")
