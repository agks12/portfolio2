import requests
from io import StringIO
import pandas as pd
import json
from dotenv import load_dotenv
import os
from datetime import datetime
from pathlib import Path
import sqlite3
import xml.etree.ElementTree as ET


root_dir = r'D:\공공데이터포털\최종결과'
load_dotenv(dotenv_path=os.path.join(root_dir,'keyinfo.env'))
SERVICE_KEY_EKAPE = os.getenv('SERVICE_KEY_EKAPE')


with open(os.path.join(root_dir, "table_config.json"), 'r', encoding="utf-8") as f:
    table_config = json.load(f)

# json 파일 필터링, file name / # key
unused_api_file = ['축산물품질평가원_계란등급판정정보', # 계란등급판정결과조회
                     '축산물품질평가원_농장식별번호정보', # 농장식별번호 정보
                     '축산물품질평가원_쇠고기이력정보', # 소 및 쇠고기 개체정보, 소 이동정보
                     '축산물품질평가원_축산물등급판정확인서 발급정보', # 가금산물 등급판정확인서 발급번호정보, 소 묶음번호 정보, 소,돼지 등급판정확인서 발급번호정보
                     '축산물품질평가원_축산물통합이력정보', # 축산물이력정보조회
                     '축산물품질평가원_축산물DNA동일성검사 업소정보',# 축산물DNA동일성검사 업소정보 조회
                     '축산물품질평가원_축산물경락가격정보', # 미사용
                     '농림축산식품부_가축 통계 데이터', # 농림축산식품부 전제 미사용중
                     '농림축산식품부_대상질병별 농장 통계 데이터 제공',
                     '농림축산식품부_사육규모별 농장 통계 데이터 제공',
                     '농림축산식품부_업종별 농장 통계 데이터 제공',
                     '농림축산식품부_운영상태별 농장 통계 데이터 제공',
                     '농림축산식품부_원천산지위판장가격목록조회 (1999-01-01 ~ 2016-01-19)',
                     '농림축산식품부_원천조사가격목록조회 (1998-01-03 ~ 2022-12-30)',
                     '농림축산식품부_지역별 농장 통계 데이터 제공',
                     '농림축산식품부_축종별 농장 통계 데이터 제공',]
                     # 여기서 부터는 이미 확보된 데이터여서 사용 안함
                     #'축산물품질평가원_가금산물 일일거래가격정보']

# 하나의 json 파일 속 여러 url 들 필터링  key / # file name
unused_api_key = ['계란 등급판정결과(확인서)정보', # 축산물등급판정정보
                  '닭 등급판정결과(확인서)정보',
                  '돼지 등급판정 부위별고기및부산물 정보',
                  '돼지 등급판정결과(확인서) 정보',
                  '말 등급판정결과(확인서)정보',
                  '소 등급판정(확인서) 부위별고기및 부산물 정보',
                  '소 등급판정결과(확인서)정보',
                  '오리 등급판정결과(확인서)정보',
                  '축산물 도매시장 소도체 등급별 경락가격정보',
                  '축산물 도매시장 소도체 1++등급 세부경락가격정보',
                  '축산물 소도체 육질육량 등급별 경락가격정보',
                  '소도체 1++ 등급 세부경락가격 상세정보',
                  '축산물 돼지도체 경락가격 상세정보',
                  '소 권역별 경락가격현황',
                  '돼지 권역별 경락가격 현황',
                  '축산물 실시간 경매시황중계정보',
                  '실시간 소도체 경매현황 조회',
                  '축산물 실시간 돼지도체 경매세부현황정보',
                  '축산물 실시간 돼지도체 등급별경매현황정보',
                  '축산물 소도체 경락상세정보',
                  '돼지도체 등급별 경락가격정보',
                  '소 도매시장 정보',
                  '축산물 실시간 소도체 등급별경매현황정보',
                  '계란 검수(소비)확인정보', # 축산물유통정보
                  '닭 검수(소비)확인정보',
                  '돼지 검수(소비)확인정보',
                  '소 검수(소비)확인정보',
                  '오리 검수(소비)확인정보']
                  # 여기서 부터는 이미 확보된 데이터여서 사용 안함
                  #'돈육대표가격 현황']

# 전체 조회 되지 않는 것들 날짜 칼럼
one_day_ymd_dict = {'돼지 권역별 경락가격 현황':{'D':'baseYmd'}, # 축산물등급판정정보
                    '실시간 소도체 경매현황 조회':{'D':'auctDate'},
                    '축산물 실시간 경매시황중계정보':{'D':'auctDate'},
                    '축산물 실시간 돼지도체 경매세부현황정보':{'D':'auctDate'},
                    '축산물 실시간 돼지도체 등급별경매현황정보':{'D':'auctDate'},
                    '축산물 실시간 소도체 등급별경매현황정보':{'D':'auctDate'},
                    '순별 축산물소비자가격 정보':{'M':'standYm'}, # 축산물유통정보
                    '연도별 축산물소비자가격 정보':{'Y':'standY'},
                    '월별 축산물소비자가격 정보':{'M':'standYm'},
                    '일자별 축산물소비자가격 정보':{'D':'standYmd'},
                    '축산물 재고동향':{'Y':'standYyyy','M':'standMm'}
                    } # 소 도매시장 정보는 날짜없음

date_col = {
    '순별 축산물소비자가격 정보':{'M':'standYm'}, # 축산물유통정보
    '연도별 축산물소비자가격 정보':{'Y':'standY'},
    '월별 축산물소비자가격 정보':{'M':'standYm'},
    '일자별 축산물소비자가격 정보':{'D':'standYmd'},
    '축산물 재고동향':{'Y':'standYyyy','M':'standMm'}

}

# 요청 파리미터 샘플 사용하는 중 int -> 4310.0 같은 포맷 에러
req_format_error = ['judgeKind','itemCd']



def read_json(save_path):
    """
    api json파일에서 정보 추출
    """
    with open(save_path,'r',encoding='utf-8') as f:
        data = json.load(f)
    url_dict = data

    sample_params = {}
    match_col = {}

    state = 'ekape'

    # mafra - 농림축산식품부, ekape - 공공데이터포털
    req_filter = {'ekape':{'item':'항목명(영문)','data':'샘플데이터'}, 'mafra':{'item':'항목명','data':'샘플데이터'}}
    res_filter = {'ekape':{'col_en':'항목명(영문)','col_ko':'항목명(국문)'},'mafra':{'col_en':'항목명','col_ko':'항목설명'}}

    for key,value in url_dict.items():

        # url 설정 - 농림축산식품부 API 는 샘플 파라미터 들어간형태
        if value['url'].split('/')[2] == '211.237.50.150:7080': # 농림축산식품부 API
            url = value['url'].split('/sample')[0]
            state = 'mafra'
        else:
            url = value['url']


        # 샘플 입력 요청 파라미터 생성
        req = pd.DataFrame(value['req']['data'],columns=value['req']['columns'])

        sample = dict(zip(req[f'{req_filter[state]['item']}'],req[f'{req_filter[state]['data']}']))#'항목명(영문)'],req['샘플데이터']))

        if 'API_KEY' in sample:
            sample.pop('API_KEY')
        # if 'startYmd' in sample:
        #     sample['startYmd'] = (datetime.today() - timedelta(days=1)).strftime("%Y%m%d")#'20000101'
        # if 'endYmd' in sample:
        #     sample['endYmd'] = (datetime.today() - timedelta(days=1)).strftime("%Y%m%d")#datetime.today().strftime("%Y%m%d")
        if "judgeKind" in sample:
            sample["judgeKind"] = "4301"


        if "itemCd" in sample:
            sample["itemCd"] = None


        sample_params[key] = {'url':url,'params':sample}

        res = pd.DataFrame(value['res']['data'],columns=value['res']['columns'])
        # 영어 - 한글 칼럼 이름
        if state == 'ekape':
            # 불필요 칼럼 제거
            remove_list = [
                'resultCode',
                'resultMsg',
                'numOfRows',
                'pageNo',
                'totalCount'
            ]
            res = res[~res['항목명(영문)'].isin(remove_list)]

        col_dict = dict(zip(res[f'{res_filter[state]['col_en']}'],res[f'{res_filter[state]['col_ko']}']))
        # 소문자 변경
        col_dict = {
            k.lower(): v
            for k, v in col_dict.items()
        }

        # 날짜 칼럼 없는 api 칼럼 임의로 추가
        if key=="월별 축산물소비자가격 정보" or key=="순별 축산물소비자가격 정보":
            col_dict['standym'] = '기준월'

        if key=="쇠고기 부분육 경락가격 조회" or key=="제주 돼지도체 등급별 경락가격 정보 조회":
            col_dict['standymd'] = '조회일자'

        if key=='일자별 축산물소비자가격 정보':
            col_dict['standymd_c'] = '조회일자'

        if key=="연도별 축산물소비자가격 정보":
            col_dict['standy'] = '기준연도'

        if key=="축산물 재고동향":
            col_dict['standyyyy'] = '기준연도'
            col_dict['standmm'] = '기준월'

        match_col[key] = {'url':url,'col':col_dict}

    return sample_params, match_col, state



def ekape_request(value):
    """
    API 호출
    """
    response = requests.get(
        value['url'],
        params={
            **value["params"],
            "serviceKey": SERVICE_KEY_EKAPE
        },
        timeout=30)
    response.raise_for_status()

    #print('진입성공 ',value)
    return response.text


def ekape_res_xml_to_pd(key,value):
    """
    api로 반환된 xml or json db에 넣기 쉽게 pd로 변환
    """
    # xml pd테이블로 변환 데이터 없으면 에러
    matched_keys = [i for i in req_format_error if i in value['params']]

    for i in matched_keys:
        #print(f'데이터 포맷 변경 {i} {value['params'][i]} -> {int(value['params'][i])}')
        value['params'][i] = int(value['params'][i])

    try:
        res_text = ekape_request(value)
        #print(res_text)
        #print(value)

        root = ET.fromstring(res_text)
        items = root.findall(".//item")

        # 데이터 없음
        if not items:
            print(f"{key} - 데이터 없음")
            return None

        # 데이터 있음
        df_tmp = pd.read_xml(
            StringIO(res_text),
            xpath="//item"
        )

        return df_tmp

    except ET.ParseError as e:
        print(f"XML 파싱 에러: {e}")
        return None

    except Exception as e:
        print(f"API 반환 에러: {e}")
        return None




def request_one_day(key, value):
    """
    각 api별 다르게 존재 하는 날짜 파라미터에 값 전달 하기 위한 함수
    """

    # 기본 startYmd, endYmd

    value_copy = value.copy()

    poi = 0

    col_year = None
    col_month = None
    col_day = None
    if key in one_day_ymd_dict:
        col_year = one_day_ymd_dict[key]['Y'] if 'Y' in one_day_ymd_dict[key] else None  # yyyy
        col_month = one_day_ymd_dict[key]['M'] if 'M' in one_day_ymd_dict[key] else None  # yyyymm / 예외 축산물 재고동향 m
        col_day = one_day_ymd_dict[key]['D'] if 'D' in one_day_ymd_dict[key] else None  # yyyymmdd

        # 연도 순회
        if col_year is not None and col_month is None: poi = 4
        # 연도 월 순회 / 월 널이 아니면 무조건 월 순회
        if col_month is not None: poi = 6
        # 연도 월 일 순회 / 일 널이 아니면 무조건 일 순회
        if col_day is not None: poi = 8

    else:
        poi = 10

    ymd = datetime.today().strftime("%Y%m%d")

    y = ymd[0:4]

    m = ymd[4:6]

    value_copy['params'][col_year] = y
    value_copy['params'][col_month] = y + m
    if key == '축산물 재고동향':
        value_copy['params'][col_year] = y
        value_copy['params'][col_month] = m

    value_copy['params'][col_day] = ymd
    if poi == 10:
        value_copy['params']['startYmd'] = ymd
        value_copy['params']['endYmd'] = ymd

    #print(key,value_copy)
    ym_df = ekape_res_xml_to_pd(key,value_copy)

    # 반환된 테이블에 임의로 만든 칼럼 강제 추가
    if key == "연도별 축산물소비자가격 정보":
        ym_df['standy'] = y

    if key == "쇠고기 부분육 경락가격 조회" or key == "제주 돼지도체 등급별 경락가격 정보 조회" or key == "일자별 축산물소비자가격 정보":
        ym_df['standymd_c'] = ymd

    if ym_df is None:return None # None 반환인 경우 리턴

    #print(ym_df)
    # 축산물 유통정보 -> 순,월별 축산물소비자 가격정보는 날짜 없어서 추가해줌
    if key != '축산물 재고동향':
        ym_df['standym'] = y + m
    return ym_df


def each_request(key, value,state):
    """
    축산물품질평가원 - 농림축산식품부 API 구조 달라서 구분하기 위한 함수
    """
    df = None

    try:
        if state=='ekape':
            if 'numOfRows' in value['params'].keys():
                value['params']['numOfRows'] = 1000

            df = request_one_day(key,value)

        # if state=='mafra':
        #
        #     #print(value['params'])
        #
        #     req_url = value['url'] # base url
        #
        #     req_url += f'/{str(SERVICE_KEY_MAFRA)}' # 인증키
        #
        #     for i,v in value['params'].items():
        #         req_url += rf"/{v}"
        #
        #     response = requests.get(
        #         req_url,
        #         timeout=30
        #     )
        #     df = pd.read_xml(StringIO(response.text), xpath="//row")

        return df
    except Exception as e:
        print(f'Request 요청 에러 {e}')
        return None

def upset_db(key,df):
    """
    DB에 오늘 일자 데이터 존재 하면 pass 없으면 insert
    """
    conn = sqlite3.connect(os.path.join(root_dir, "data.db"))
    cursor = conn.cursor()

    db_table_name = table_config[key]['table_name']  # 국문 시트명으로 DB테이블 명 찾기
    db_table_col = table_config[key]['columns'].keys()

    df = df[db_table_col]  # 엑셀에서 읽은 테이블 칼럼 순서 변경

    try:
        tmp_data = list(
            df.itertuples(index=False, name=None)
        )

        select_query = f"""
            SELECT 1
            FROM {db_table_name}
            WHERE {" AND ".join([f"{col} IS ?" for col in db_table_col])}
            LIMIT 1
        """

        insert_query = f"""
            INSERT INTO {db_table_name} ({", ".join(db_table_col)})
            VALUES ({", ".join(["?"] * len(db_table_col))})
        """

        inserted_rows = []
        skipped_rows = []

        for row in tmp_data:

            check = cursor.execute(
                select_query,
                row
            ).fetchone()

            if check is None:
                cursor.execute(
                    insert_query,
                    row
                )
                inserted_rows.append(row)

            else:
                skipped_rows.append(row)

        conn.commit()

        print(f"\n[{db_table_name}]")
        print(f"INSERT 된 데이터: {len(inserted_rows)}건")
        print(f"이미 존재해서 제외된 데이터: {len(skipped_rows)}건")

        print("\n--- INSERT 된 데이터 ---")
        for row in inserted_rows:
            print(row)

        print("\n--- 이미 존재한 데이터 ---")
        for row in skipped_rows:
            print(row)
    except Exception as e:
        print(f'에러발생 {key} {e}')

    # 저장 확정
    conn.commit()

    # 연결 종료
    conn.close()


def livestock_grading_information(save_path):

    json_dict = read_json(save_path)

    url_dict = json_dict[0]
    col_dict = json_dict[1]
    state = json_dict[2]

    for key, value in url_dict.items():
        if any(i in key for i in unused_api_key): continue # 사용하지 않는 상세 api 건너띄기
        try:
            df = each_request(key,value, state)

            if df is None:
                continue

            df.columns = df.columns.str.lower()

            for col in col_dict[key]["col"]:
                if col not in df.columns:
                    df[col] = None

            df = df[list(col_dict[key]["col"].keys())]

            upset_db(key,df)

        except Exception as e:
            print(f'에러발생 {e}')



def search_json():
    """
    경로에 api 정보 관련 json 파일 순회 하면서 하나씩 읽기
    json 파일에 url, 입출력 파라미터 정보 존재
    """
    for file in Path(os.path.join(root_dir)).rglob("*.json"):
        if '~$' not in file.name and 'url_data' in file.name :
            if any(i in file.name for i in unused_api_file) :continue

            livestock_grading_information(file)



if __name__=="__main__":
    search_json()
