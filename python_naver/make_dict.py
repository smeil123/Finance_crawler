import json
import os
import re
from konlpy.tag import Twitter

tagger = Twitter() # Twitter 태깅 함수
# 불용어 리스트 (본인의 판단에 따라 더 추가 가능) : 텍스트 분석에 필요가 없거나 사업보고서 특성상 당연히 많이 나와 분석에 의미 없는 단어들
stop_words = ['사항','제기','및','년','사업','관','그','등','것','및','부','수','위','나','대하','개월','원']

def read_text(text,flag,pos_neg_dict):
    # 파일 읽고 dictionary로 만듦
    f = open(text, 'r', encoding = 'UTF-8')
    dict_text = f.read()
    f.close()

    dict_text = dict_text.split('\n')
    dict_text[0] = re.sub("\ufeff",'',dict_text[0])

    words = [tagger.nouns(doc) for doc in dict_text if len(tagger.nouns(doc))!=0]

    # text를 dict형태로 변환
    for word in words:
        for w in word:
            w = re.findall(r'[가-힣]{2,10}', w)
            for p in w:
                if p is None: continue
                if p not in stop_words:
                    pos_neg_dict[p] = flag
    return pos_neg_dict

if __name__ == "__main__":
    print("pos_pol_word, neg_pol_word 파일명이어야합니다.")
    textpath = input("textfile위치(n - 현재폴더) : ")

    if textpath == 'n':
        textpath = './'

    pos_neg_dict = {}
    pos_neg_dict = read_text(os.path.join(textpath,"pos_pol_word.txt"),'pos',pos_neg_dict)
    pos_neg_dict = read_text(os.path.join(textpath,"neg_pol_word.txt"),'neg',pos_neg_dict)

    # dict를 json으로 저장
    json = json.dumps(pos_neg_dict)
    f = open("dict.json","w")
    f.write(json)
    f.close()

    print("dictionary가 json으로 저장되었습니다")
