import pandas as pd 
import os
from datetime import datetime

manual = """무엇이든 상상할 수 있는 사람은 무엇이든 만들어 낼 수 있다 - 앨런 튜링

- 최초 작성일 : 2024.3.20
- 버전 : 1.0
- 작성자 : C1U0137

- 기본 인터페이스 명령
          0 : 매뉴얼 다시 보기
          1 : 파일 업로드
          2 : 파일 업데이트(수정)
          3 : 종료
          
1. 업로드 파일은 반드시 csv파일로 하여 정해진 업로드 폴더 경로에 두고 명령을 실행할 것
2. 업로드 파일의 이름 형식을 반드시 지킬 것
    - 통합공정번호_vp_datasheet_tag_작성자이름_yymmdd(날짜).csv
    - 대소문자 역시 구분함
    예시 : 2204_vp_datasheet_tag_이상현_240320.csv
3. 업로드 데이터의 형식이 지켜지지 않을 경우 파일에 있는 데이터가 모두 올라가지 않음
4. 분할로 여러 번 올리는 것 가능함
5. 파일 업데이트는 허락을 득한 후 암호키를 받고 실행
"""

base_path = os.chdir()
metadata_path = ""
upload_folder_path = ""
merged_process = pd.read_csv("병합공정 정보.csv")
writer_list = ["서영빈",
                   "이가은",
                   "김민영",
                   "지승환",
                   "이지승",
                   "변재아",
                   "강여원",
                   "박채림",
                   "정채원",]

def check_file_name(input_file_name) :
    merged_process = input_file_name[:4]
    vp_datasheet_tag = input_file_name[4:22]
    writer_name = input_file_name[22:25]
    upload_date = input_file_name[26:32]
    is_csv = input_file_name[-4:]

    error_msg = "파일 이름 오류. 업로드 파일 제목 형식 다시 확인"
    
    date = today()

    try :
        if int(merged_process) in process_merged_list :
            check_flag = True
    except :
        print(error_msg)
        return False
    
    if vp_datasheet_tag != "_vp_datasheet_tag_" :
        print(error_msg)
        return False
    
    if writer_name not in writer_list :
        print(error_msg)
        return False
    
    if upload_date != date :
        print("업로드 날짜와 오늘 날짜 불일치")
        return False
    
    if is_csv != '.csv' :
        print("csv 파일 아님")
        return False
    
    return True


def upload_data(file_name) :

    merged_process = file_name[:4]

    try :
        new_df = pd.read_csv(os.path.join(upload_folder_path, file_name), encoding='cp949')
    except :
        try :
            new_df = pd.read_csv(os.path.join(upload_folder_path, file_name), encoding='uft8')
        except :
            print("업로드 오류 : 폴더 경로 다시 확인")
    
    try :
        df = pd.read_csv(os.join.path(), encoding='cp949')
    except :
        try :
            df = pd.read_csv(os.join.path(), encoding='utf8')
        except :
            print("메인 파일 불러오기 실패. 관리자에게 문의")

    chk_dupl_key = check_duplicate_key(new_df)
    chk_dupl_key_in_df = check_duplicate_key_in_df(new_df, df)
    chk_process_no = check_process_number(new_df)
    chk_blank_tagNo = check_blank_tagNo(new_df)
    chk_right_cct = check_right_cct(new_df)
    chk_blank_cct = check_blank_cct(new_df)

    print("업로드 완료")


if __name__ == '__main__' :
    print(manual)
    while True :
        input_cmd = input("실행번호 (0/1/2/3) : ")
        
        if input_cmd == "0" :
            print(manual)
        
        elif input_cmd == "1" :
            input_file_name = input("업로드 파일 이름 입력(확장자까지)")
            chk_filename = check_file_name(input_file_name)
            upload_data(file_name)

            if chk_filename :
                upload_data(chk_filename)
            else :
                continue

        elif input_cmd == "2" :
            print(manual)

        elif input_cmd == "3" :
            break
