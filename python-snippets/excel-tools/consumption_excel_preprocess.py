import pandas as pd
import os
import re
import numpy as np
import sys
from openpyxl import load_workbook
from collections import defaultdict
import openpyxl
from openpyxl.styles import Alignment
import warnings
import traceback

warnings.filterwarnings('ignore') # 경고 무시

#######################################################################################################
# 가구 데이터 변수
#######################################################################################################
# 현재 파일 경로
BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) \
    else os.path.dirname(os.path.abspath(__file__))

# 가구 데이터 년월별 excel 파일 경로
FUR_ROOT_DIR = os.path.join(BASE_DIR, "가구_총소비량(24,25, 26년)")

# py 실행시 하드코등 경로 적용
#FUR_ROOT_DIR = r'D:\소비데이터_전체_원본\0. 자료 구성\가구_총소비량(24,25, 26년)'

# 채널별_특성별_구매량_구매액_총소비량 결과 파일 이름'
char_channel_file_name = '채널별_특성별_구매량_구매액_총소비량_1차_결과_테이블형식.xlsx'
# 세부품목별_구매량_구매액_총소비량 결과 파일 이름
detail_crop_value_file_name = '세부품목별_구매량_구매액_총소비량_1차_결과_테이블형식.xlsx'
# 가구 엑셀 가공본 파일 이름
fur_process_file_name = '전체별_구매량_구매액_총소비량_엑셀가공본.xlsx'

# 가구 데이터 키값
char_channel_key_cols = ['연도', '월', '품목', '값지표', '통계지표']
detail_crop_value_key_cols = ['연도', '월', '품목', '값지표', '통계지표', '세부품목']

# 가구 데이터 구분 # 특성별, 세부품목별, 채널별 - 구매량, 구매액, 총소비량 = 3*3=9가지 경우
value_types = ['구매량', '총소비량', '구매액']
group_types = ['채널별', '특성별', '세부품목별']

# 가구 데이터 특성별 칼럼 정지포인트
columns_break_point = {
    ('채널별', '구매량'): '지인구매',
    ('채널별', '구매액'): '지인구매',
    ('채널별', '총소비량'): '직접재배',
    ('특성별', '구매량'): '600만이상',
    ('특성별', '구매액'): '600만이상',
    ('특성별', '총소비량'): '600만이상'}


#######################################################################################################
# pos 데이터 변수
#######################################################################################################
# pos데이터 원본 경로
POS_ROOT_DIR = os.path.join(BASE_DIR, "POS 데이터","농산물_POS_3월 리포트_20260415","1_월별 리포트")
# py 실행시 하드코등 경로 적용
#POS_ROOT_DIR = r'D:\소비데이터_전체_원본\0. 자료 구성\POS 데이터\농산물_POS_3월 리포트_20260415\1_월별 리포트'

# pos데이터 가공본 파일 이름
pos_process_file_name = '전체_pos데이터_엑셀가공본.xlsx'

# pos데이터 결과 파일 이름
pos_result_file_name = '전체_pos데이터_테이블형식.xlsx'
pos_result_key_cols = ['연도', '월', '품목']

# cpi 주민등록통계 정보 파일 -
CPI_RESI_ROOT_DIR = os.path.join(BASE_DIR, "POS 데이터")
# py 실행시 하드코등 경로 적용
#CPI_RESI_ROOT_DIR = r'D:\소비데이터_전체_원본\0. 자료 구성\POS 데이터'

# cpi 경로 & 주민 통계 경로
cpi_desi_file = os.path.join(CPI_RESI_ROOT_DIR,'POS 데이터.xlsx')

# 가구 데이터 품목 리스트 - 양배추 다음 배추, 건고추 다음 고추 와야함  , 가공란, 고추, 당근, 양배추 안씀
CROP_NM = ['가공란','감자','건고추','계란','고추','김치',
           '당근','마늘','무','양배추','배추','사과','쌀','양파']

# 가구 데이터 최종 테이블 칼럼
POS_TABLE_COL = [
    'CPI',
    '오프라인_총판매액(천원)','오프라인_총판매액(천원,실질)','오프라인_총판매액(억 원)',
    '오프라인_총판매액(억원,실질)','가격(원/kg)','중량(톤)',
    '2페이지_대형마트','2페이지_기업형슈퍼','2페이지_체인슈퍼','2페이지_조합마트',
    '2페이지_편의점','2페이지_개인',
    '대형마트_온라인판매액','대형마트_온라인비중',
    '지역별_서울','지역별_경기,인천,강원','지역별_부산,경남','지역별_대구,경북',
    '지역별_대전,세종,충정','지역별_광주,전라,제주'
]

# 가구 데이터 시트 이름 리스트
pos_sheet_list = ['판매액','판매량','평균가']


#######################################################################################################
# 전체 통합 테이블 변수
#######################################################################################################
FINAL_TABLE_COL = [
    #'연도','월','품목',
    '소비정보분석사업_가구소비량', # 특성별, 채널별 구매량,구매액,총소비량에서 가져옴
    '소비정보분석사업_가구구매량', # 특성별, 채널별 구매량,구매액,총소비량에서 가져옴
    '소비정보분석사업_가구구매액', # 특성별, 채널별 구매량,구매액,총소비량에서 가져옴
    '소비정보분석사업_제조업사용량',
    '소비정보분석사업_외식업사용량',
    '소비정보분석사업_급식업사용량',
    '소비정보분석사업_군급식사용량',
    '소비정보분석사업_학교급식사용량',
    'POS_판매량','POS_판매단가', # POS 가격/중량에서 가져옴
    '가계동향조사_가구당월평균구매액'
]

# pos데이터 가공본 파일 이름
final_process_file_name = '최종_통합_결과_엑셀가공본.xlsx'

# pos데이터 결과 파일 이름
final_result_file_name = '최종_통합_결과_테이블형식.xlsx'
final_result_key_cols = ['연도', '월', '품목']



# 멀티 칼럼으로 만드는 함수 / 처음 두줄
def multi_col_merge(header_df, break_col):
    # 3. 결측치(NaN)를 앞의 값으로 채우기 (병합된 셀 처리용) # 가로 채우기임
    header_df = header_df.ffill(axis=1)

    # 4. 1층과 2층의 텍스트를 언더바(_)로 합치기
    new_columns = []
    for col_idx in range(header_df.shape[1]):
        level1 = str(header_df.iloc[0, col_idx]).strip()
        level2 = str(header_df.iloc[1, col_idx]).strip()

        # 'nan' 문자열 예외 처리 및 병합
        if level1 == 'nan' or level1 == level2:
            col_name = level2
        elif level2 == 'nan':
            col_name = level1
        else:
            col_name = f"{level1}_{level2}"

        new_columns.append(col_name)

        # 별별로 특정 칼럼까지만 사용함
        if level2 == break_col:
            break
    return new_columns


def fur_data_walk_merge():
    # 딕셔너리 별로 테이블 저장하기 위한 함수
    group_tables = defaultdict(list) # 채널 특성별
    detail_crop_table = pd.DataFrame() # 세부 품목별

    for root, dirs, files in os.walk(FUR_ROOT_DIR):  # 하위 파일 전부 순회

        for f in files:
            # xlsx파일만 읽음
            if not f.endswith(".xlsx") or f.startswith("~$"): continue

            # 구매량, 구매액, 총소비량 중 어떤거 매치되는지 찾기
            matched_value_fime = next((c for c in value_types if c in f), None)
            # 특성별, 세부품목별, 채널별 중 어떤거 매치되는지 찾기
            matched_group_fime = next((c for c in group_types if c in f), None)

            # 특성별, 세부품목별, 채널별 - 구매량, 구매액, 총소비량 만족하는 경우만 읽음
            if matched_value_fime and matched_group_fime:
                path = os.path.join(root, f)  # 파일 전체 경로 이름
                # print(matched_value_fime,matched_group_fime)
                now_excel = pd.read_excel(path, header=None)

                # 무조건 뒤에서부터 분리해서 연도 월 찾음
                year = path.split('\\')[-3][0:4]  # 폴더 이름으로 년도 구하기
                month = path.split('\\')[-2]  # 폴더 이름으로 월 구하기
                print(f'{year} {month} {matched_group_fime} {matched_value_fime} 파일 처리중')

                # 202504 기준 특성별 구매량 구매액 총소비량 구조 같음 - 제일 처음한 코드 구조
                # 202504 기준 세부품목별 구매량 구매액 총소비량 구조 같음 -
                # 202504 기준 채널별 구매량 구매액 구조 같음 총소비량 칼럼 아래 비중 줄 추가(2504만) 및 무상 칼럼 있음

                # 특성별, 채널별 과 세부품목별은 구조 달라서 다른 함수 적용

                # 세부품목별인 경우
                if matched_group_fime == '세부품목별':
                    # 1. 첫 번째로 NaN 아닌 행 찾기 '품목인 경우만 해당 2025년4월 같은거 버려야함'
                    header_idx = now_excel[now_excel.iloc[:, 0] == '품목'].index[0]

                    # 첫번쨰 칼럼(품목이름) 채우기
                    now_excel.iloc[:, 0] = now_excel.iloc[:, 0].ffill()

                    # 2. 무조건 품목 시점부터 아래 2번째 칸부터 사용
                    header_df = now_excel.iloc[header_idx + 2]
                    now_excel.columns = header_df
                    now_excel = now_excel.iloc[header_idx + 2:].reset_index(drop=True)

                    # 2번쨰 칼럼이 nan인 부분부터 아래 제거 - 아래 낙서있으면 있으면 널값 채워짐
                    now_excel = now_excel.loc[:now_excel[now_excel.iloc[:, 1].isna()].index.min() - 1]

                    # 첫번쨰 칼럼(품목이름) 채우기
                    now_excel.iloc[:, 0] = now_excel.iloc[:, 0].ffill()

                    # 테이블 값에 칼럼이 들어가 있는 형태이므로 각각 테이블로 분리 필요 - 첫번째 칼럼 값
                    groups = now_excel.groupby(now_excel.iloc[:, 0])

                    result = []
                    for name, g in groups:

                        # 세부품목 row (첫 row)
                        header = g.iloc[0, 2:]

                        # 데이터 row들
                        data = g.iloc[1:, :]

                        # 품목 하나씩 순회
                        for i in range(len(data)):
                            row = data.iloc[i]

                            temp = pd.DataFrame({
                                '품목': name,
                                '통계지표': row.iloc[1],
                                '세부품목': header,
                                '값': row.iloc[2:].values
                            })

                            temp.insert(0, '연도', year)
                            temp.insert(1, '월', month)
                            temp.insert(2, '값지표', matched_value_fime)

                            temp = temp[temp['통계지표'] != '비중']  # 비중 row 버림
                            temp = temp[~temp['세부품목'].isna()]  # 세부품목 nan 버림

                            result.append(temp)

                    # 엑셀 파일 하나
                    final = pd.concat(result, ignore_index=True)

                    # 전체 테이블에 합침
                    detail_crop_table = pd.concat([detail_crop_table, final], ignore_index=True)

                # 특성별, 채널별인 경우
                else:
                    # 1. 첫 번째로 NaN 아닌 행 찾기 '품목인 경우만 해당 2025년4월 같은거 버려야함'
                    header_idx = now_excel[now_excel.iloc[:, 0] == '품목'].index[0]

                    # 2. '품목' 행과 그 바로 아래행을 슬라이싱하여 멀티 칼럼 정보로 가져오기
                    header_df = now_excel.iloc[header_idx: header_idx + 2]

                    # 5. 원래 데이터프레임에 새 칼럼 적용하고
                    # 칼럼 분할 지점 설정
                    break_col = columns_break_point[(matched_group_fime, matched_value_fime)]
                    # 새로운 멀티 칼럼
                    new_columns = multi_col_merge(header_df, break_col)
                    # 기존 테이블에서 필터링된 칼럼 아닌거 버림
                    now_excel = now_excel.iloc[:, :len(new_columns)]
                    # 칼럼 적용
                    now_excel.columns = new_columns

                    # 2번쨰 칼럼은 빈값이라 "품목" 으로 적용됨
                    now_excel.columns.values[1] = '통계지표'

                    # 2번쨰 칼럼이 nan인거 제거 - 아래 통계 구하는거 있으면 널값 채워짐
                    # now_excel = now_excel.dropna(subset=[now_excel.columns[1]])

                    # 3. 그 아래만 데이터로 사용
                    now_excel = now_excel.iloc[header_idx + 2:].reset_index(drop=True)

                    # 품목 칼럼 채우기
                    now_excel['품목'] = now_excel['품목'].ffill()

                # 연도 월 칼럼 추가
                now_excel.insert(0, '연도', year)
                now_excel.insert(1, '월', month)
                now_excel.insert(2, '그룹지표', matched_group_fime)
                now_excel.insert(3, '값지표', matched_value_fime)

                # 전체를 하나의 table에 그룹별로 붙이기
                group_tables[(matched_group_fime, matched_value_fime)].append(now_excel)

    # 채널,특성별(group_tables) ,세부품목별(detail_crop_table) 총 2가지 반환
    return group_tables, detail_crop_table #


# 후처리 함수
def fur_data_postprocessing(char_channel_table, detail_crop_table):


    # 최종 특성,채널별 테이블 머지
    # 그룹테이블 하나의 딕셔너리 테이블로
    final_tables = {
        k: pd.concat(v, ignore_index=True)
        for k, v in char_channel_table.items()
    }

    # 각 지표 별 merge -- 특성별에만 가구당 평균값 있어서 특성별 기준 merge
    result_amount = pd.merge(final_tables['특성별', '구매량'], final_tables['채널별', '구매량'],
                             on=char_channel_key_cols, how='left')
    result_quantity = pd.merge(final_tables['특성별', '구매액'], final_tables['채널별', '구매액'],
                               on=char_channel_key_cols, how='left')
    result_consumption = pd.merge(final_tables['특성별', '총소비량'], final_tables['채널별', '총소비량'],
                                  on=char_channel_key_cols, how='left')
    # 전체 세로 합치기(concat)
    char_channel = pd.concat([result_amount, result_quantity, result_consumption]).drop(columns=['그룹지표_x', '그룹지표_y'])

    # 2025년 4월거는 2배 해야함 15일치만 나와서
    mask = (
            (char_channel['연도'] == '2025') &
            (char_channel['월'] == '4')
    )
    # 제외 칼럼(문자열 칼럼)
    num_cols = char_channel.columns.difference(char_channel_key_cols)
    char_channel.loc[mask, num_cols] *= 2

    # 최종 세부품목별 테이블
    # 5월 추정량(톤) 은 추정량(톤)이랑 같음
    detail_crop_table['통계지표'] = detail_crop_table['통계지표'].str.replace('5월 ', '', regex=False)
    # 비중 칼럼 계산
    group_sum = detail_crop_table.groupby(
        char_channel_key_cols
    )['값'].transform('sum')
    denom = group_sum.astype(float).replace(0, np.nan)
    detail_crop_table['비중'] = (
            detail_crop_table['값'].astype(float)
            .div(denom)
            .fillna(100)
            .astype(float)
            * 100
    )

    # 2025년 4월거는 2배 해야함 15일치만 나와서
    mask = (
            (detail_crop_table['연도'] == '2025') &
            (detail_crop_table['월'] == '4')
    )
    # 제외 칼럼(문자열 칼럼)
    num_cols = detail_crop_table.columns.difference(detail_crop_value_key_cols)
    detail_crop_table.loc[mask, num_cols] *= 2


    # 가구 데이터 품목 가공김치 -> 김치, 고추 -> 건고추로 변경
    char_channel['품목'] = char_channel['품목'].replace({
        '가공김치': '김치',
        '고추': '건고추',
    })
    detail_crop_table['품목'] = detail_crop_table['품목'].replace({
        '가공김치': '김치',
        '고추': '건고추',
    })

    # 월 형식 MM으로 통일
    char_channel['월'] = char_channel['월'].astype(int).astype(str).str.zfill(2)  # 월 형식 MM
    detail_crop_table['월'] = detail_crop_table['월'].astype(int).astype(str).str.zfill(2)  # 월 형식 MM

    # 채널,특성별(group_tables) ,세부품목별(detail_crop_table) 총 2가지 반환
    return char_channel, detail_crop_table


# 최종 테이블 형식 파일 저장
# 만약 이미 파일 존재하면 연도, 월, 품목, 통계지표 보고 존재하면 pass 없으면 채우기 해야함
# upset - 이미 있으면 업데이트 , 없으면 생성
def check_upset_create_file(current_df, file_path, key_col):
    if os.path.exists(file_path):


        # 이미 존재하는 파일 불러오기 -- 새로 생기는 경우
        old_df = pd.read_excel(file_path,dtype=str)

        # 인덱스 매칭 안되는 경우 없애기 위한 월 형식 통일
        current_df['월'] = current_df['월'].astype(int).astype(str).str.zfill(2) # 월 형식 MM
        old_df['월'] = old_df['월'].astype(int).astype(str).str.zfill(2) # 월 형식 MM


        # 키 칼럼 전부 str타입 변경
        for col in key_col:
            old_df[col] = old_df[col].astype(str).str.strip()
            current_df[col] = current_df[col].astype(str).str.strip()

        # 키 칼럼 아닌것들 숫자형 타입 변경
        num_cols = old_df.columns.difference(key_col)
        old_df[num_cols] = old_df[num_cols].astype(float)
        current_df[num_cols] = current_df[num_cols].astype(float)

        # 엑셀 파일 인덱스 설정
        old_df = old_df.set_index(key_col).sort_index()

        # 새로 만든 테이블 인덱스 생성
        new_df = current_df.set_index(key_col).sort_index()

        # 신규 행 -- 키값 없는것만 신규
        inserted = new_df.loc[
            ~new_df.index.isin(old_df.index)
        ]

        # 업데이트 -- 키값 있는거는 업데이트
        old_df.update(new_df)


        # 공통 인덱스
        common_idx = old_df.index.intersection(new_df.index)
        updated_count = len(common_idx)

        # 변경된 값 비교

        # 최종 결과
        result = pd.concat([
            old_df,
            inserted
        ]).reset_index()

        inserted_count = len(inserted)

        result.to_excel(file_path, index=False)

        print(f"INSERT : {inserted_count}행")
        print(f"UPDATE : {updated_count}행")
        print(f'{file_path}파일 있음 upset')
    else:
        # 파일 없음
        current_df.to_excel(file_path, index=False)
        print(f'{file_path}파일 없음 create')


def fur_extract_excel_view(char_channel_table,detail_crop_table):
    # 엑셀 보기좋게 작업
    # 이건 어차피 최종 파일을 작업하는거라

    pro1 = char_channel_table[
        (char_channel_table['통계지표'] == '구매량 (톤)') |  # 구매량
        (char_channel_table['통계지표'] == '구매액 (억원)') |  # 구매액
        (char_channel_table['통계지표'] == '추정량 (톤)')]  # 총소비량

    # 멀티 칼럼 적용 / '_' 으로 분리
    new_cols = []
    for col in pro1.columns:
        if '_' in col:
            new_cols.append(tuple(col.split('_', 1)))
        else:
            new_cols.append(('', col))

    pro1.columns = pd.MultiIndex.from_tuples(new_cols)

    groups = pro1.groupby(('', '품목'))

    with pd.ExcelWriter(fur_process_file_name, engine='openpyxl') as writer:
        for key, value in groups:
            #print(key)
            v_copy = value.copy()

            v_copy = v_copy.drop(columns=[('', '값지표'), ('', '품목')])

            # 세부품목별 지표 병합하기

            # 필요한 품목만 세부품목에서 필터링
            filtered = detail_crop_table[
                detail_crop_table['품목'] == key
                ]

            # 피벗으로 세부품목 칼럼 만들기
            df_wide = filtered.pivot_table(
                index=['연도', '월', '값지표', '품목', '통계지표'],
                columns='세부품목',
                values='값',
                aggfunc='sum'
            ).reset_index()

            df_wide = df_wide.drop(columns=['품목', '값지표'])  # 필요할 때만

            # 병합하려면 같은 멀티 인덱스여야 해서 멑티인덱스 적용
            df_wide.columns = pd.MultiIndex.from_tuples([
                ('', col) if col in ['연도', '월', '통계지표']
                else ('세부품목별', col)
                for col in df_wide.columns
            ])

            # 타입 맟추기
            for col in [('', '연도'), ('', '월'), ('', '통계지표')]:
                v_copy[col] = v_copy[col].astype(str)
                df_wide[col] = df_wide[col].astype(str)

            # 병합
            merge_result = v_copy.merge(
                df_wide,
                on=[('', '연도'), ('', '월'), ('', '통계지표')],
                how='left'
            )

            # 정렬을 위한 타입 변경
            merge_result[('', '월')] = merge_result[('', '월')].astype(int)

            # 멀티 인덱스 설정(인덱스 병합되서 보기 편함)
            merge_result = merge_result.set_index([
                ('', '연도'),
                ('', '월'),
                ('', '통계지표')
            ])

            # 정렬
            merge_result = merge_result.sort_values([
                ('', '연도'),
                ('', '월'),
                ('', '통계지표')
            ])

            merge_result.index.names = ['연도', '월', '통계지표']

            # 시트명 = 품목 이름
            sheet_name = str(key)  # 엑셀 제한 (31자)
            merge_result.to_excel(writer, sheet_name=sheet_name)

    # 2. [완벽 수정] 상위 기준을 고려한 조건부 세로 셀 병합
    wb = openpyxl.load_workbook(fur_process_file_name)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # [헤더 정리] A, B, C열 1~2행 병합 및 제목 세팅
        headers = {'A': '연도', 'B': '월', 'C': '통계지표'}
        for col_letter, title in headers.items():
            ws[f'{col_letter}1'] = title
            ws[f'{col_letter}2'] = None
            ws.merge_cells(f'{col_letter}1:{col_letter}2')
            ws[f'{col_letter}1'].alignment = Alignment(horizontal='center', vertical='center')
        # 3번째 행을 숨기기 -- 삭제하면 인덱스 없어져서 그런가 아래 병합 꺠짐
        ws.row_dimensions[3].hidden = True

    # 파일 저장
    wb.save(fur_process_file_name)

    print(f"가구 데이터 엑셀 가공본 저장 {fur_process_file_name}")


#######################################################################################################
# pos 데이터 함수 시작
#######################################################################################################
# CPI 정보 가져오기
def CPI_info():
    cpi = pd.read_excel(cpi_desi_file, sheet_name='CPI raw', dtype='str')

    # 칼럼 공백 제거 & 테이블 전체(문자열) 값에 공백제거  / cpi시트 전체 공백존재함
    cpi.columns = cpi.columns.str.strip()
    cpi = cpi.apply(lambda col: col.astype(str).str.strip() if col.dtype == "object" else col)

    # 연도 월 칼럼 추가
    cpi.insert(0, '연도', cpi['시점'].str.split('.').str[0])
    cpi.insert(1, '월', cpi['시점'].str.split('.').str[1])

    # 이후 매칭하기 하는데서 인덱스로 매칭할므로 인덱스 설정
    cpi = cpi.set_index(['연도', '월'])

    cpi = cpi.rename(columns={'고춧가루': '건고추', '달걀': '계란'})

    return cpi


# 주민등록통계 정보 가져오기
def resident_info():
    # 칼럼 없어야되서 header = None적용
    resident = pd.read_excel(cpi_desi_file, sheet_name='주민등록통계', dtype='str', header=None)

    header_df = resident.iloc[0:2]

    # 멀티 칼럼 적용
    resident.columns = multi_col_merge(header_df,[])

    # 멀티칼럼 적용전 칼럼 버리기
    resident = resident.iloc[2:].reset_index(drop=True)

    # 칼럼 이름 중에 '자동'이라는 글자가 '없는' 칼럼들만 선택
    clean_columns = [col for col in resident.columns if '자동' not in col]

    # 해당 칼럼들만 남겨서 df를 업데이트합니다.
    resident = resident[clean_columns]

    # 연도 월 칼럼 추가
    resident.insert(0, '연도', resident['시점'].str.split('.').str[0])
    resident.insert(1, '월', resident['시점'].str.split('.').str[1])

    # 칼럼 공백 제거 & 테이블 전체(문자열) 값에 공백제거  / cpi시트 전체 공백존재함
    resident.columns = resident.columns.str.strip()
    resident = resident.apply(lambda col: col.astype(str).str.strip() if col.dtype == "object" else col)

    # 값 없는 칼럼 제거
    resident = resident.dropna(axis=1)
    # 중복된 칼럼 중 첫 번째만 남기고 나머지는 제거하기 - 멀티칼럼 적용하면서 중복칼럼생김
    resident = resident.loc[:, ~resident.columns.duplicated(keep='first')]

    # 1. 서울 (단일 칼럼 숫자 변환)
    resident['서울'] = pd.to_numeric(resident['서울특별시_계 (명)'], errors='coerce')

    # 2. 경기_인천_강원
    cols_gyeonggi = ['경기도_계 (명)', '인천광역시_계 (명)', '강원도_계 (명)']
    resident['경기_인천_강원'] = resident[cols_gyeonggi].apply(pd.to_numeric, errors='coerce').sum(axis=1)

    # 3. 부산_경남
    cols_busan = ['부산광역시_계 (명)', '울산광역시_계 (명)', '경상남도_계 (명)']
    resident['부산_경남'] = resident[cols_busan].apply(pd.to_numeric, errors='coerce').sum(axis=1)

    # 4. 대구_경북
    cols_daegu = ['대구광역시_계 (명)', '경상북도_계 (명)']
    resident['대구_경북'] = resident[cols_daegu].apply(pd.to_numeric, errors='coerce').sum(axis=1)

    # 5. 대전_충정
    cols_daejeon = ['대전광역시_계 (명)', '충청남도_계 (명)', '충청북도_계 (명)', '세종특별자치시_계 (명)']
    resident['대전_충정'] = resident[cols_daejeon].apply(pd.to_numeric, errors='coerce').sum(axis=1)

    # 6. 광주_전라_제주
    cols_gwangju = ['광주광역시_계 (명)', '전라남도_계 (명)', '전라북도_계 (명)', '제주특별자치도_계 (명)']
    resident['광주_전라_제주'] = resident[cols_gwangju].apply(pd.to_numeric, errors='coerce').sum(axis=1)

    # 필요 칼럼만 필터링
    resident = resident[['연도', '월', '서울', '경기_인천_강원', '부산_경남', '대구_경북', '대전_충정', '광주_전라_제주']]

    # 이후 매칭하기 하는데서 인덱스로 매칭할므로 인덱스 설정
    resident = resident.set_index(['연도', '월'])

    return resident


def pos_data_walk_merge():
    # 품목별, 시트별 원본 테이블 저장
    origin_dict = {}

    # 폴더 순회하면서 품목별 파일 모두 읽어서 딕셔너리에 시트별 테이블로 저장
    for root, dirs, files in os.walk(POS_ROOT_DIR):
        for f in files:
            # xlsx파일만 읽음 / 열려있는거 안읽음
            if not f.endswith(".xlsx") or f.startswith("~$"): continue
            path = os.path.join(root, f)  # 파일 경로

            # 어떤 품목 파일이 사용되는지 확인 - 정해진(CROP_NM) 품목만 사용됨
            matched_crop = next((c for c in CROP_NM if c in f), None)
            if matched_crop:
                #print(matched_crop)
                # 시트 3번 반복
                for sheet_name in pos_sheet_list:
                    print(f'{matched_crop} {sheet_name} 파일 작업중')
                    # 멀티 칼럼 사용하기 위해 헤더 칼럼 적용안함
                    now_excel = pd.read_excel(path, sheet_name=sheet_name, header=None)

                    # 무조건 첫번째 두번쨰 줄을 칼럼으로 사용
                    header_df = now_excel.iloc[0:2]

                    # 멀티 칼럼 적용
                    now_excel.columns = multi_col_merge(header_df,[])

                    # 멀티칼럼 적용전 칼럼 버리기
                    now_excel = now_excel.iloc[2:].reset_index(drop=True)

                    # 19-Jan 처럼 날짜 형식 되있어서 1월19일로 인식하지만 실제 19년1월임 / 파싱필요
                    now_excel['MONTH'] = now_excel['MONTH'].astype(str)
                    year = '20' + now_excel['MONTH'].str.split('-').str[2].str[:2]
                    month = now_excel['MONTH'].str.split('-').str[1]
                    now_excel.insert(0, '연도', year)
                    now_excel.insert(1, '월', month)
                    # 품목 칼럼 추가하기 - 하나의 테이블로 만들때 - 칼럼 달라서 사용안함
                    # now_excel['품목'] = matched_crop
                    # origin_table = pd.concat([origin_table,now_excel], ignore_index=True)

                    # print(now_excel)
                    # 키 없으면 에러 나서 setdefault 사용
                    origin_dict.setdefault(matched_crop, {})[sheet_name] = now_excel
    # 품목별, 시트별 원본 테이블 반환
    return origin_dict


def pos_data_cal_match(origin_dict):
    # cpi정보 테이블
    cpi = CPI_info()
    resident = resident_info()
    # 최종 pos결과 품목별 테이블 저장
    result_dict = {}
    for crop, sheets in origin_dict.items():
        #print(crop)
        # 사용 안함
        if crop in ('가공란', '양배추', '당근', '고추'): continue

        # 원본 건드리지 않기 위해 복사본 생성
        amount = sheets['판매액'].copy()
        volume = sheets['판매량'].copy()
        price = sheets['평균가'].copy()

        # 월 형식 통일
        amount['월'] = amount['월'].astype(int).astype(str).str.zfill(2)  # 월 형식 MM
        volume['월'] = volume['월'].astype(int).astype(str).str.zfill(2)  # 월 형식 MM
        price['월'] = price['월'].astype(int).astype(str).str.zfill(2)  # 월 형식 MM

        # 인덱스 생성 - 고유값
        amount = amount.set_index(['연도', '월'])
        volume = volume.set_index(['연도', '월'])
        price = price.set_index(['연도', '월'])



        # 최종 테이블도 매칭하기 위해 sales와 같은 인덱스 적용
        df = pd.DataFrame(columns=POS_TABLE_COL).reindex(amount.index)

        # cpi테이블에서 품목에 해당하는 열만 가져옴 - 이때 숫자로 변환, cpi시트 읽을때 문자열로 읽음
        df['CPI'] = pd.to_numeric(cpi[crop], errors='coerce')

        # 칼럼 계산적용 -- 연도,월 을 기준으로 매칭함
        df['오프라인_총판매액(천원)'] = amount['오프라인 전채널_총판매액(천원)']

        df['오프라인_총판매액(천원,실질)'] = amount['오프라인 전채널_총판매액(천원)'] * (100 / df['CPI'])

        df['오프라인_총판매액(억 원)'] = amount['오프라인 전채널_총판매액(천원)'] / 100000

        df['오프라인_총판매액(억원,실질)'] = df['오프라인_총판매액(천원,실질)'] / 100000

        if crop == '계란':
            df['가격(원/kg)'] = price['오프라인 전채널_가중평균가격(원/구)']

            df['중량(톤)'] = ((volume['오프라인 전채널_총판매량(천구)_추정_하한']
                            + volume['오프라인 전채널_총판매량(천구)_추정_상한']) / 2) / 100000
        else:
            df['가격(원/kg)'] = price['오프라인 전채널_가중평균가격(원/kg)']

            df['중량(톤)'] = ((volume['오프라인 전채널_총판매량(kg)_추정_하한']
                            + volume['오프라인 전채널_총판매량(kg)_추정_상한']) / 2) / 1000

        df['2페이지_대형마트'] = amount['1.대형마트_총판매액(천원)'] / 100000

        df['2페이지_기업형슈퍼'] = amount['2.기업형슈퍼_총판매액(천원)'] / 100000

        df['2페이지_체인슈퍼'] = amount['2-1.체인슈퍼_총판매액(천원)'] / 100000

        df['2페이지_조합마트'] = amount['2-2.조합마트_총판매액(천원)'] / 100000

        df['2페이지_편의점'] = amount['3.편의점_총판매액(천원)'] / 100000

        df['2페이지_개인'] = amount['4.개인슈퍼_총판매액(천원)'] / 100000

        df['대형마트_온라인판매액'] = amount['온라인_대형마트_총판매액(천원)'] / 100000

        df['대형마트_온라인비중'] = df['대형마트_온라인판매액'] / (df['대형마트_온라인판매액'] + df['2페이지_대형마트']) * 100

        df['지역별_서울'] = amount['오프라인 지역별_서울'] / resident['서울'] * 1000

        df['지역별_경기,인천,강원'] = (amount['오프라인 지역별_경기'] +
                              amount['오프라인 지역별_인천'] +
                              amount['오프라인 지역별_강원']) / resident['경기_인천_강원'] * 1000

        df['지역별_부산,경남'] = (amount['오프라인 지역별_부산'] +
                           amount['오프라인 지역별_울산'] +
                           amount['오프라인 지역별_경남']) / resident['부산_경남'] * 1000

        df['지역별_대구,경북'] = (amount['오프라인 지역별_대구'] +
                           amount['오프라인 지역별_경북']) / resident['대구_경북'] * 1000

        df['지역별_대전,세종,충정'] = (amount['오프라인 지역별_대전'] +
                              amount['오프라인 지역별_충남'] +
                              amount['오프라인 지역별_충북']) / resident['대전_충정'] * 1000

        df['지역별_광주,전라,제주'] = (amount['오프라인 지역별_광주'] +
                              amount['오프라인 지역별_전남'] +
                              amount['오프라인 지역별_전북'] +
                              amount['오프라인 지역별_제주']) / resident['광주_전라_제주'] * 1000
        result_dict[crop] = df


    # 모든 품목 한 테이블로 합치기
    result_origin_table = pd.DataFrame()
    for i, v in result_dict.items():
        v_copy = v.copy()
        v_copy = v_copy.reset_index()
        v_copy.insert(2, '품목', i)
        result_origin_table = pd.concat([result_origin_table, v_copy])
    result_origin_table.sort_values(['연도', '월', '품목'])

    # 월 형식 통일
    result_origin_table['월'] = result_origin_table['월'].astype(int).astype(str).str.zfill(2)  # 월 형식 MM

    # 품목별 딕셔너리, 모두 하나로 합친 한 테이블 반환
    return result_dict, result_origin_table


def pos_extract_excel_view(result_dict):
    with pd.ExcelWriter(pos_process_file_name, engine='openpyxl') as writer:
        for i, v in result_dict.items():

            v_copy = v.copy()

            new_cols = []
            for col in v_copy.columns:
                if '_' in col:
                    new_cols.append(tuple(col.split('_', 1)))
                else:
                    new_cols.append(('', col))

            v_copy.columns = pd.MultiIndex.from_tuples(new_cols)
            # 정렬을 위한 타입 변경

            v_copy.index.names = ['연도', '월']

            v_copy.to_excel(writer, sheet_name=str(i))

    wb = openpyxl.load_workbook(pos_process_file_name)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # [헤더 정리] A, B, C열 1~2행 병합 및 제목 세팅
        headers = {'A': '연도', 'B': '월'}
        for col_letter, title in headers.items():
            ws[f'{col_letter}1'] = title
            ws[f'{col_letter}2'] = None
            ws.merge_cells(f'{col_letter}1:{col_letter}2')
            ws[f'{col_letter}1'].alignment = Alignment(horizontal='center', vertical='center')
        # 3번째 행을 숨기기 -- 삭제하면 인덱스 없어져서 그런가 아래 병합 꺠짐
        ws.row_dimensions[3].hidden = True

    # 파일 저장
    wb.save(pos_process_file_name)
    print(f"pos 데이터 엑셀 가공본 저장 {pos_process_file_name}")

#######################################################################################################
# 전체 통합 테이블 함수 시작
#######################################################################################################
def final_integration(con_info,pos_info):
    # 총 합계 계산

    # 읍면/동 여부별_읍면 + 읍면/동 여부별_동 가 총 합
    con_info_copy = con_info.copy()
    con_info_copy['가구총합'] = (
            con_info_copy['읍면/동 여부별_읍면'].fillna(0) +
            con_info_copy['읍면/동 여부별_동'].fillna(0)
    )

    # 칼럼,row 필터링
    con_info_copy = con_info_copy[['연도', '월', '품목', '통계지표', '가구총합']]
    con_info_copy = con_info_copy[
        (con_info_copy['통계지표'] == '구매량 (톤)') |
        (con_info_copy['통계지표'] == '구매액 (억원)') |
        (con_info_copy['통계지표'] == '추정량 (톤)')]

    # 피벗으로 통계지표 칼럼으로 만들기
    con_info_copy = con_info_copy.pivot(
        index=final_result_key_cols,
        columns='통계지표',
        values='가구총합'
    )

    # pos테이블 인덱스 없는데 인덱스 설정
    pos_info = pos_info.set_index(final_result_key_cols)

    # 각각 인덱스 추출
    con_idx = con_info_copy.index
    pos_idx = pos_info.index

    # 인덱스 합집합 추출 - 양쪽에 존재하는 모든 인덱스 사용
    all_idx = con_idx.union(pos_idx)

    # 각 테이블 합집합 인덱스 적용
    con_info_copy = con_info_copy.reindex(all_idx)
    pos_info = pos_info.reindex(all_idx)

    # 최종 테이블도 매칭하기 위해 sales와 같은 인덱스 적용
    df = pd.DataFrame(columns=FINAL_TABLE_COL).reindex(all_idx)

    df['POS_판매량'] = pos_info['가격(원/kg)']
    df['POS_판매단가'] = pos_info['중량(톤)']
    df['소비정보분석사업_가구소비량'] = con_info_copy['추정량 (톤)']
    df['소비정보분석사업_가구구매량'] = con_info_copy['구매량 (톤)']
    df['소비정보분석사업_가구구매액'] = con_info_copy['구매액 (억원)']

    final_table = df.reset_index()
    #df_origin.to_excel('최종통계_원본.xlsx', index=False)

    # 품목row 칼럼으로 올림 -- excel 보기 편함 용도
    excel_view = df.unstack('품목')
    excel_view.columns = excel_view.columns.swaplevel(0, 1)
    excel_view = excel_view.sort_index(axis=1, level=0)

    # 현재 2중 칼럼인데 3중 칼럼으로 만들기
    new_cols = []
    for lev0, lev1 in excel_view.columns:

        if '_' in lev1:
            part1 = lev1.split('_')[0]
            part2 = lev1.split('_')[1]
        else:
            part1 = ''
            part2 = lev1

        new_cols.append((lev0, part1, part2))

    excel_view.columns = pd.MultiIndex.from_tuples(new_cols)

    # 엑셀 작업용(멀티 인덱스), 원본 반환
    return excel_view, final_table

def final_extract_excel_view(excel_view):
    with pd.ExcelWriter(final_process_file_name, engine='openpyxl') as writer:
        excel_view.to_excel(writer, sheet_name='view')

    wb = openpyxl.load_workbook(final_process_file_name)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # [헤더 정리] A, B, C열 1~2행 병합 및 제목 세팅
        headers = {'A': '연도', 'B': '월'}
        for col_letter, title in headers.items():
            ws[f'{col_letter}1'] = title
            ws[f'{col_letter}2'] = None
            ws.merge_cells(f'{col_letter}1:{col_letter}3')
            ws[f'{col_letter}1'].alignment = Alignment(horizontal='center', vertical='center')
        # 3번째 행을 숨기기 -- 삭제하면 인덱스 없어져서 그런가 아래 병합 꺠짐
        ws.row_dimensions[4].hidden = True

    # 파일 저장
    wb.save(final_process_file_name)
    print(f"통합 데이터 엑셀 가공본 저장 {final_process_file_name}")

if __name__ == "__main__":
    try:

        #######################################################################################################
        # 가구 데이터 작업 시작
        #######################################################################################################
        # 가구 데이터 엑셀 파일 순회 하면서 합치기

        print('가구 데이터 엑셀 파일 순회 시작')
        fur_char_channel_table, fur_detail_crop_table = fur_data_walk_merge()

        # 후처리 하기
        fur_char_channel_table, fur_detail_crop_table = fur_data_postprocessing(fur_char_channel_table, fur_detail_crop_table)
        print('가구 데이터 병합 완료')

        # 특성별, 채널별 파일 존재 여부에 따른 upset, create
        check_upset_create_file(fur_char_channel_table, char_channel_file_name, char_channel_key_cols)
        # 세부품목별 파일 존재 여부에 따른 upset, create
        check_upset_create_file(fur_detail_crop_table, detail_crop_value_file_name, detail_crop_value_key_cols)

        # 가구 데이터 엑셀 보기 좋게 만들기
        fur_extract_excel_view(fur_char_channel_table,fur_detail_crop_table)

        #######################################################################################################
        # pos 데이터 작업 시작
        #######################################################################################################
        print('\npos 데이터 엑셀 파일 순회 시작')
        pos_origin_dict = pos_data_walk_merge()

        crop_group_dict, pos_result_merge_table = pos_data_cal_match(pos_origin_dict)
        print('pos 데이터 병합 완료')

        # pos 최종 테이블 파일 존재 여부에 따른 upset, create
        check_upset_create_file(pos_result_merge_table, pos_result_file_name, pos_result_key_cols)

        # pos 데이터 엑셀 보기 좋게 만들기
        pos_extract_excel_view(crop_group_dict)

        #######################################################################################################
        # 전체 통합 테이블 작업 시작
        #######################################################################################################
        print('\n전체 통합 파일 작업 시작')
        excel_view_table, final_origin_table = final_integration(fur_char_channel_table,pos_result_merge_table)

        # 최종 테이블 파일 존재 여부에 따른 upset, create
        check_upset_create_file(final_origin_table, final_result_file_name, final_result_key_cols)

        # 엑셀용
        final_extract_excel_view(excel_view_table)

        print('\n모든 작업이 완료되었습니다.')

    except Exception as e:
        print(f'\n오류 발생: {e}')
        traceback.print_exc()

    finally:
        input('\n엔터를 누르면 종료됩니다...')
