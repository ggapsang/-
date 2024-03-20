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


# 리스트 안에 리스트 풀기

def flatten_list(nested_list):
    return [item for sublist in nested_list for item in sublist]


# 리스트 길이 맞추기

def pad_list_to_length(original_list, target_length):
    # 리스트 길이가 목표 길이보다 작은 경우, 차이만큼 None을 추가
    while len(original_list) < target_length:
        original_list.append(None)
    return original_list


# 딕셔너리 key : value(list)를 df로 변환

import pandas as pd

# pad_list_to_length 함수를 정의합니다.
def pad_list_to_length(original_list, target_length):
    while len(original_list) < target_length:
        original_list.append(None)
    return original_list

# 딕셔너리 리스트 예시
dict_list = [
    {'key': 'A', 'values': pad_list_to_length([1, 2, 3], 4)},
    {'key': 'B', 'values': pad_list_to_length([4, 5], 4)},
    {'key': 'C', 'values': pad_list_to_length([6, 7, 8, 9], 4)}
]

# DataFrame을 생성하기 위해 딕셔너리를 변환합니다.
# 'key' 열에는 각 키가, 그리고 'values'를 나누어 각각의 value1, value2, ... 열에 매핑합니다.
data_for_df = {'key': [item['key'] for item in dict_list]}
for i in range(max(len(d['values']) for d in dict_list)):  # 최대 길이에 맞게 반복
    data_for_df[f'value{i+1}'] = [d['values'][i] if i < len(d['values']) else None for d in dict_list]

# pandas DataFrame을 생성합니다.
df = pd.DataFrame(data_for_df)

# DataFrame을 출력합니다.
print(df)


