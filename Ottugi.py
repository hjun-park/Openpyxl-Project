# -*- coding:utf-8 -*-

import sys
from datetime import datetime
import openpyxl
import os
import time

'''     
    # 엑셀 객체 생성
    wb = openpyxl.load_workbook('file.xlsx')
    
    # 새로운 시트 생성
    wb.create_sheet(title=None, index=None)
    
    # 워크 시트 선택
    sheet = wb.active   # 현재 활성 중인 워크시트
    sheet = wb["sheetname"] # sheetname이라는 워크시트 선택
    
    # 워크시트 목록 조회
    wb.get_sheet_names()
    
    # 특정 쉘 불러오기
    C1 = sheet[C1]
    
    # 쉘 객체에서 row/column , 좌표자체, 저장된 값 얻을 수 있음
    C1.row  / C1.column  / C1.coordinate / C1.value
    
    # 쉘에 데이터 입력 방법
    sheet.cell(row=1, column=1).value = 10

    # 엑셀 쉘 범위 접근 방법
      특정 범위 접근 : cell_range = sheet[‘A1’:’C2’]
      특정 row 접근 : row2 = sheet[2]
      특정 row 범위 : row_range = sheet[5:10]
      특정 Column : colC = sheet[‘C’]
      특정 Column 범위 : col_range = sheet[‘C:D’]
      
    # 쉘 가운데 정렬
    sheet.cell(row=current_row, column=1).alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

    # 시트 최대 행 열 불러오기
    number_row = curr_sheet.max_row number_col = curr_sheet.max_column

'''

# 열 너비 자동 맞춤
# culumns is passed by list and element of columns means column index in worksheet.
# if culumns = [1, 3, 4] then, 1st, 3th, 4th columns are applied autofit culumn.
# margin is additional space of autofit column.

def AutoFitColumnSize(worksheet, columns=None, margin=2):
    for i, column_cells in enumerate(worksheet.columns):
        is_ok = False
        if columns == None:
            is_ok = True
        elif isinstance(columns, list) and i in columns:
            is_ok = True

        if is_ok:
            length = max(len(str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = length + margin

    return worksheet


def make_xls_file(filepath):
    filepath = "/test.xlsx"
    wb = openpyxl.Workbook()
    wb.save(filepath)


def access_shell_range():
    wb = openpyxl.load_workbook('file.xlsx')
    sheet = wb["sheetname"]
    allList = []
    for row in sheet.iter_rows(min_row=1, max_row=10, min_col=2, max_col=5):
        a = []
        for cell in row:
            a.append(cell.value)
        allList.append(a)




def init_file(filepath):
    wb = openpyxl.Workbook()

    # 시트 내용 삭제
    for sheet in wb.sheetnames:
        wb.remove(wb[sheet])

    ws = wb.create_sheet(title='2021', index=0)  # 워크시트 이름 설정


def read_from_file(filepath):
    wb = openpyxl.load_workbook(filename=filepath)
    ws = wb.active
    # ws = wb['2021']   # 워크시트 이름으로 접근가ㅡㄴㅇ

    # min_col : 실제 데이터가 존재하는 곳부터 시작
    for column in ws.iter_cols(min_col=3, max_col=ws.max_column):
        val_list = []
        for cell in column:
            value = cell.value
            val_list.append('{:>15}'.format(value))

        print(''.join(val_list))
    wb.close()


if __name__ == '__main__':
    # 0. 결과파일을 저장할 파일 입력
    dirpath = "D:\\Ottugi"
    temp = input("Saved result file: ")
    result_file_name = os.path.join(dirpath, temp)
    wb_master = openpyxl.load_workbook(result_file_name)
    ws_master = wb_master.active

    # 1. 파일 리스트 출력
    print("=============== Directory Check ================")
    if os.path.isdir(dirpath) is False:
        print(f'{dirpath} doesn\'t exist. Check it')
        sys.exit(0)
    else:
        print(f"[{dirpath}] OK")

    file_list = [os.path.join(dirpath, f) for f in os.listdir(dirpath) if os.path.isfile(os.path.join(dirpath, f))]

    file_list.remove(result_file_name)

    for num, file in enumerate(file_list):
        print(f'\t\t[{num+1}] ==> {file}')

    # while True:
    ANSWER = input("\n\nContinue ? (Y/N) ")
    if ANSWER == "n" or ANSWER == "N":
        sys.exit(0)
    elif ANSWER != "y" or ANSWER != "Y":
        pass # break
    else:
        pass
    # else:
    #     # continue    : 다시 입력받도록


    # 1.5 정보 입력받기
    while True:
        column_match_list = []
        print("== 빈 값이 입력되면 진행됩니다 ==")
        print("복사할 행 / 붙여넣을 행 순서대로 입력")
        print("예시 : B A")

        while True:
            column_info = list(map(str, input().split()))

            if not column_info:
                break

            column_match_list.append(column_info)

        print("=========== 컬럼 정보 ===========")
        for column_info in column_match_list:
            print(f'[{column_info[0]}] ==> [{column_info[1]}]')

        ANSWER = input("\n\nContinue ? (Y/N) ")
        if ANSWER == "n" or ANSWER == "N":
            continue

        elif ANSWER != "y" or ANSWER != "Y":
            break
        else:
            continue


    # 2. 행은



    # 2. 엑셀 열기
    for file in file_list:
        wb_slave = openpyxl.load_workbook(file)
        ws_slave = wb_slave.active # wb["sheetname"]

        # 특정 셀 복사
        m_row = ws_slave.max_row
        m_col = ws_slave.max_column

        print(f'[{file}] ==> {m_row}/{m_col}')

        # 특정 컬럼에 대해서 row가 몇 개까지 있는지 확인
        A1 = ws_slave["A1"]

        print(A1.value)

        max_row_for_c = max((c.row for c in ws['C'] if c.value is not None))






        # for i in range(1, m_row + 1):
        # # for j in range(1, m_col + 1):
        #     c = ws_slave.cell(row = i, column= 4)   # 열 고정, 행 변화
        #     ws_master.cell(row = i, column = 4).value = c.value


    # wb_master.save(str(result_file_name))
