import sys
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import XMLConverter, HTMLConverter, TextConverter
from pdfminer.layout import LAParams
import io
import os

##  프로그레스바
import sys
def progressBar(value, endvalue, bar_length=20):
    percent = float(value) / endvalue
    arrow = '-' * int(round(percent * bar_length)-1) + '>'
    spaces = ' ' * (bar_length - len(arrow))

    sys.stdout.write("\rPercent: [{0}] {1}%".format(arrow + spaces, int(round(percent * 100))))
    sys.stdout.flush()


# - data : pdf파일
# - txtfp : 읽지못한 pdf목록을 open한 것

def pdfparser(data,txtfp):

    fp = open(data, 'rb')
    rsrcmgr = PDFResourceManager()
    retstr = io.StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    # Create a PDF interpreter object.
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    # Process each page contained in the document.
    filename = data
    data = ""
    for i,page in enumerate(PDFPage.get_pages(fp)):
        try:
            interpreter.process_page(page)
            data =  retstr.getvalue()
        except:
            print(filename+"파일 %d번째 페이지 읽지 못함"% (i+1))
            txtfp.write(filename+"파일 %d번째 페이지 읽지 못함"% (i+1))
            txtfp.write("\n")
            pass
    return data

# - dirname : pdf파일 경로
# - savefolder : pdf to txt를 저장할 경로

def pdfread(dirname,savefolder=None):

    # 해달 폴더에 있는 파일명 가져옴
    file_list = os.listdir(dirname)
    path_filenames = [dirname+'/'+i for i in file_list]

    # text파일 저장할 폴더 생성
    if savefolder is None:
        down_path = "./textfolder"
    else:
        down_path = savefolder
    try:
        os.mkdir(down_path)
    except:
        print("savefolder 있음")

    # 확장자명을 제외한 파일명만 추출
    text_list = os.listdir(down_path)
    text_list = [os.path.splitext(i)[0] for i in text_list]

    print("총 %d개의 파일 존재"%len(file_list))

    f = io.open(os.path.join("읽지못한pdf"+".txt"),'w',encoding='utf8')

    for i,file in enumerate(file_list):
        if file in text_list:
            continue

        # pdf 파일만 read
        if '.pdf' == os.path.splitext(file)[-1]:
            progressBar(i,len(file_list))
            try:
                pdf2txt = pdfparser(path_filenames[i],f)
                if pdf2txt == "":
                    f.write("%s 읽지못함"%file)
                    f.write('\n')
                    print("%s 읽지못함"%file)
                    continue

            except:
                f.write("%s 읽지못함"%file)
                f.write('\n')
                print("%s 읽지못함"%file)
                continue

            try:
                #txt파일 저장
                with io.open(os.path.join(down_path,file+".txt"),'w',encoding='utf8') as pdf_f:
                    pdf_f.write(pdf2txt)
                pdf_f.close()
            except:
                f.write('%s를 txt파일 저장 실패'%file)
                f.write('\n')
                print('%s를 txt파일 저장 실패'%file)

    f.close()

if __name__ == "__main__":
    pdfpath = input("pdf 위치 : ")
    downpath = input("txt 저장받을 위치( n - ./textfolder): ")

    if downpath == 'n':
        downpath = None

    # ex) ./temp, ./temp_text
    pdftext = pdfread(pdfpath,downpath)
    if pdftext is not None:
        with io.open(os.path.join("읽지못한pdf"+".txt"),'w',encoding='utf8') as f:
            for pt in pdftext:
                f.write(pt)
                f.write("\n")
            f.close()
