##### PDF 주석에서 텍스트 추출

import fitz

def extract_annotations_from_pdf(pdf_path) :
    doc = fitz.open(pdf_path)
    annotations = []

    for page in doc :
        annots = page.annots()
        if annots :
            for annot in annots :
                info = annot.info
                annotations.append(info['content'])

    doc.close()

    return annotations

# PyMuPDF 모듈 설치
# pdf 파일에서 주석 정보를 가져옴
# pdf_path : str, 경로와 파일명 및 확장자가 모두 포함된 pdf 파일


##### PDF 파일 자르기

import PyPDF2
import os

def cut_pdf_pages_tail(pdf_path, temp_pdf_path, rng_pages) : # 잘랐을 때의 마지막 pdf 페이지
    with open(pdf_path, 'rb') as pdf_file :
        reader = PyPDF2.PdfReader(pdf_file)
        writer = PyPDF2.PdfWriter()

        for page_num in range(rng_pages) :
            if page_num < len(reader.pages) :
                writer.add_page(reader.pages[page_num])

        with open(temp_pdf_path, 'wb') as temp_pdf_file :
            writer.write(temp_pdf_path)

    os.remove(pdf_path)
    os.rename(temp_pdf_path, pdf_path)

# 앞에서부터 원하는 페이지 만큼 잘린 pdf파일을 임시 파일에 저장했다가 기존 파일로 이름을 바꾼 뒤 기존 파일을 삭제함
# 경로가 일치해야 잘린 pdf 파일이 기존 파일 안에 있음



##### 엑셀 파일의 pdf 변환

import win32com.client
import os

def print_to_pdf_from_excel(excel_file_path, pdf_file_path) :
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    try :
        workbook = excel.Workbooks.Open(excel_file_path)
        workbook.ExportAsFixedFormat(0, pdf_file_path) # 0은 파일 형식으로 pdf 형식을 의미함
    
    finally :
        workbook.Close(False)
        excel.Quit()

    del excel

# DRM이 걸린 파일에서는 작동하지 않을 수도 있음
# 엑셀이 설치된 윈도우 환경에서만 사용 가능



##### 파이썬 기초 트릭


# 리스트 정렬

lst = ["banana", "apple", "cherry", "blueberry"]
lst.sort(key=len, reverse=True)  # 문자열 길이에 따라
lst = sorted(lst, key=lambda x : (len(x), x)) # 문자열 길이와 알파벳 순서
lst = sorted(lst, key=lambda x : (-len(x), x)) # 문자열 길이가 내림차순일 경우


# Trim

s = " 516-A-1201-UXB1-LET"
s = s.strip()
