import pandas as pd
import sqlite3
import os
from pathlib import Path
import json

# API원본 테이블 및 PIVOT테이블 (코드테이블 제외) 모든 테이블 정보 딕셔너리
table_config = {
    "돈육대표가격 현황": {
        "table_name": "pig_representative_price",
        "columns": {
            "sumymd": "집계일",
            "sablegubn": "돈피구분",
            "costamt": "돈피지수",
        }
    },

    "축산물 재고동향": {
        "table_name": "livestock_product_inventory_trends",
        "columns": {
            "standym": "날짜 구분",
            "judgekindnm": "축종명",
            "judgekind": "축종 코드",
            "livestockpart_1": "소 돼지 안심",
            "livestockpart_2": "소 돼지 등심",
            "livestockpart_3": "소 채끝 돼지 전지",
            "livestockpart_4": "소 목심 돼지 후지",
            "livestockpart_5": "소 앞다리 돼지 삼겹살",
            "livestockpart_6": "소 우둔 돼지 목심",
            "livestockpart_7": "소 설도 돼지 갈비",
            "livestockpart_8": "소 사태 돼지 특수부위",
            "livestockpart_9": "소 갈비 돼지 잡육",
            "livestockpart_10": "소 양지",
            "livestockpart_11": "소 특수부위",
            "livestockpart_12": "소 잡육",
            "totstock": "재고량 합계",
            "unit": "단위",
            "standyyyy":"기준연도",
            "standmm":"기준월"
        }
    },

    "월별 축산물소비자가격 정보": {
        "table_name": "monthly_livestock_product_prices",
        "columns": {
            "standym": "기준월",
            "judgekindnm": "축종명",
            "itemnm": "품목명",
            "itemcd": "품목코드",
            "grdnm": "구분명",
            "unit": "단위",
            "ntslprc": "가격",
            "avgyearprc": "평년가격",
        }
    },

    "순별 축산물소비자가격 정보": {
        "table_name": "interval_10day_livestock_product_prices",
        "columns": {
            "standym": "기준월",
            "judgekind": "축종 코드",
            "judgekindnm": "축종명",
            "itemnm": "품목명",
            "itemcd": "품목코드",
            "grdnm": "구분명",
            "unit": "단위",
            "ntslprc_1": "초순가격",
            "ntslprc_2": "중순가격",
            "ntslprc_3": "하순가격",
            "ntslprc": "월평균가격",
            "avgyearprc_1": "평년 초순가격",
            "avgyearprc_2": "평년 중순가격",
            "avgyearprc_3": "평년 하순가격",
            "avgyearprc": "평년 월평균가격",
        }
    },

    "연도별 축산물소비자가격 정보": {
        "table_name": "annual_livestock_product_prices",
        "columns": {
            "year": "연도",
            "judgekind": "축종 코드",
            "judgekindnm": "축종명",
            "itemnm": "품목명",
            "itemcd": "품목코드",
            "grdnm": "구분명",
            "unit": "단위",
            "ntslprc": "가격",
            "maxprc": "최대가격",
            "minprc": "최소가격",
            "standy": "기준연도"
        }
    },

    "일자별 축산물소비자가격 정보": {
        "table_name": "daily_livestock_product_prices",
        "columns": {
            "standymd": "날짜구분(날짜/평년)",
            "judgekind": "축종 코드",
            "judgekindnm": "축종명",
            "itemnm": "품목명",
            "itemcd": "품목코드",
            "grdnm": "구분명",
            "unit": "단위",
            "ntslprc": "평균가격",
            "maxprc": "최대가격",
            "minprc": "최소가격",
            "regionprc1": "서울가격",
            "regionprc2": "부산가격",
            "regionprc3": "대구가격",
            "regionprc4": "인천가격",
            "regionprc5": "광주가격",
            "regionprc6": "대전가격",
            "regionprc7": "울산가격",
            "regionprc8": "세종가격",
            "regionprc9": "경기가격",
            "regionprc10": "강원가격",
            "regionprc11": "충북가격",
            "regionprc12": "충남가격",
            "regionprc13": "전북가격",
            "regionprc14": "전남가격",
            "regionprc15": "경북가격",
            "regionprc16": "경남가격",
            "regionprc17": "제주가격",
            "standymd_c":"조회일자"
        }
    },

    "토종닭 가격조회": {
        "table_name": "korean_chickens_prices",
        "columns": {
            "modymd": "발표일",
            "typename": "거래종류",
            "agency": "대리점 가격",
            "mart": "대규모점포 가격",
            "franchise": "프랜차이즈 가격",
            "average": "전체 가격",
        }
    },

    "산란계(노계) 가격조회": {
        "table_name": "laying_hens_spent_hens_prices",
        "columns": {
            "stddate": "기준일자",
            "stymd": "시작기준일",
            "edymd": "종료기준일",
            "price": "산란계(노계) 가격",
        }
    },

    "산란계(병아리) 가격조회": {
        "table_name": "laying_hens_chicks_prices",
        "columns": {
            "stdmonth": "기준월",
            "price": "산란계(병아리) 가격",
        }
    },

    "육계(산지) 일일가격조회": {
        "table_name": "broilers_farm_gate_prices",
        "columns": {
            "modymd": "발표일",
            "typename": "거래종류",
            "big": "육계 대 가격",
            "medium": "육계 중 가격",
            "small": "육계 소 가격",
        }
    },

    "육계(도매) 일일가격조회": {
        "table_name": "broilers_wholesale_prices",
        "columns": {
            "modymd": "발표일",
            "typename": "거래종류",
            "franchise": "육계 프렌차이즈 가격",
            "mart": "육계 마트 가격",
            "agency": "육계 대리점 가격",
            "avg": "육계 평균 가격",
        }
    },

    "계란 일일가격조회": {
        "table_name": "egg_prices",
        "columns": {
            "modymd": "발표일",
            "typename": "거래종류",
            "verybig": "계란 왕 가격",
            "special": "계란 특 가격",
            "big": "계란 대 가격",
            "medium": "계란 중 가격",
            "small": "계란 소 가격",
        }
    },

    "오리 일일가격조회": {
        "table_name": "duck_prices",
        "columns": {
            "modymd": "발표일",
            "typename": "거래종류",
            "verybig": "오리(20~26호) 가격",
        }
    },

}


# 테이블 생성 / 이름 규칙 세부URL
CREATE_TABLE_SQL_QUERY = """
-- 기본 코드 테이블

    -- 축종 코드 테이블
    DROP TABLE IF EXISTS judgekind_code;
    CREATE TABLE IF NOT EXISTS judgekind_code (
        judgekindcd TEXT NOT NULL PRIMARY KEY, -- 축종 코드
        judgekindnm TEXT NOT NULL -- 축종 명
    );
    
    --  하위축종(삼겹, 앞다리) 코드 테이블
    DROP TABLE IF EXISTS item_code;
    CREATE TABLE IF NOT EXISTS item_code (
        judgekindcd TEXT NOT NULL, -- 축종 코드
        itemcd TEXT NOT NULL, -- 품목 코드
        itemnm TEXT NOT NULL, -- 품목 명
        PRIMARY KEY (judgekindcd, itemcd)
    );
    
    --  구분(등급/원산지) 코드 테이블
    DROP TABLE IF EXISTS grd_code;
    CREATE TABLE IF NOT EXISTS grd_code (
        grdcd TEXT NOT NULL PRIMARY KEY, -- 구분 코드
        grdnm TEXT NOT NULL -- 구분 명
    );
    
    --  단위 코드 테이블
    DROP TABLE IF EXISTS unit_code;
    CREATE TABLE IF NOT EXISTS unit_code (
        unitcd TEXT NOT NULL PRIMARY KEY, -- 단위 코드
        unitnm TEXT NOT NULL -- 단위 명
    );

-- 기본 코드 테이블 INSERT
    -- 축종 코드 테이블
    INSERT INTO judgekind_code (judgekindcd,judgekindnm) VALUES ('4301','소');
    INSERT INTO judgekind_code (judgekindcd,judgekindnm) VALUES ('4304','돼지');
    
    -- 하위축종 코드 테이블
        -- 소
    INSERT INTO item_code (judgekindcd,itemcd,itemnm) VALUES ('4301','안심','21');
    INSERT INTO item_code (judgekindcd,itemcd,itemnm) VALUES ('4301','등심','22');
    INSERT INTO item_code (judgekindcd,itemcd,itemnm) VALUES ('4301','설도','36');
    INSERT INTO item_code (judgekindcd,itemcd,itemnm) VALUES ('4301','양지','40');
    INSERT INTO item_code (judgekindcd,itemcd,itemnm) VALUES ('4301','갈비','50');
        -- 돼지
    INSERT INTO item_code (judgekindcd,itemcd,itemnm) VALUES ('4304','앞다리','25');
    INSERT INTO item_code (judgekindcd,itemcd,itemnm) VALUES ('4304','삼겹살','27');
    INSERT INTO item_code (judgekindcd,itemcd,itemnm) VALUES ('4304','갈비','28');
    INSERT INTO item_code (judgekindcd,itemcd,itemnm) VALUES ('4304','목심','68');
    
    -- 등급 코드 테이블 
    INSERT INTO grd_code (grdcd,grdnm) VALUES ('g01','1등급');
    INSERT INTO grd_code (grdcd,grdnm) VALUES ('g02','1+등급');
    
    -- 단위 코드 테이블 
    INSERT INTO unit_code (unitcd,unitnm) VALUES ('u01','kg');
    INSERT INTO unit_code (unitcd,unitnm) VALUES ('u02','ton');
    INSERT INTO unit_code (unitcd,unitnm) VALUES ('u03','원/100g');


-- API원본 테이블
-- 등급판정정보 
    -- 돈육대표가격현황
    DROP TABLE IF EXISTS pig_representative_price;

    CREATE TABLE IF NOT EXISTS pig_representative_price ( 
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sumymd TEXT NOT NULL, -- 집계일
        sablegubn TEXT NOT NULL, -- 돈피구분
        costamt REAL NOT NULL -- 돈피지수
);

-- 가금산물 일일거래가격정보 
    -- 토종닭 가격
    DROP TABLE IF EXISTS korean_chickens_prices;
    CREATE TABLE IF NOT EXISTS korean_chickens_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        modymd TEXT NOT NULL, -- 발표일
        typename TEXT NOT NULL, -- 거래종류
        agency REAL NULL, -- 대리점 가격
        mart REAL NULL, -- 대규모점포 가격
        franchise REAL NULL, -- 프랜차이즈 가격
        average REAL NULL -- 전체 가격
    );
    -- 산란계(노계) 가격조회
    DROP TABLE IF EXISTS laying_hens_spent_hens_prices;
    CREATE TABLE IF NOT EXISTS laying_hens_spent_hens_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stddate TEXT NOT NULL, -- 기준일자
        stymd TEXT NOT NULL, -- 시작기준일
        edymd TEXT NOT NULL, -- 종료기준일
        price REAL NULL -- 산란계(노계) 가격
    );
    -- 산란계(병아리) 가격조회
    DROP TABLE IF EXISTS laying_hens_chicks_prices;
    CREATE TABLE IF NOT EXISTS laying_hens_chicks_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stdmonth TEXT NOT NULL, -- 기준월
        price REAL NULL -- 산란계(병아리) 가격
    );
    -- 육계(산지) 일일가격조회
    DROP TABLE IF EXISTS broilers_farm_gate_prices;
    CREATE TABLE IF NOT EXISTS broilers_farm_gate_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        modymd TEXT NOT NULL, -- 발표일
        typename TEXT NOT NULL, -- 거래종류
        big REAL NULL, -- 육계 대 가격
        medium REAL NULL, -- 육계 중 가격
        small REAL NULL -- 육계 소 가격
    );
    -- 육계(도매) 일일가격조회
    DROP TABLE IF EXISTS broilers_wholesale_prices;
    CREATE TABLE IF NOT EXISTS broilers_wholesale_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        modymd TEXT NOT NULL, -- 발표일
        typename TEXT NOT NULL, -- 거래종류
        franchise REAL NULL, -- 육계 프렌차이즈 가격
        mart REAL NULL, -- 육계 마트 가격
        agency REAL NULL, -- 육계 대리점 가격
        avg REAL NULL -- 육계 평균 가격
    );
    -- 계란 일일가격조회
    DROP TABLE IF EXISTS egg_prices;
    CREATE TABLE IF NOT EXISTS egg_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        modymd TEXT NOT NULL, -- 발표일
        typename TEXT NOT NULL, -- 거래종류
        verybig REAL NULL, -- 계란 왕 가격
        special REAL NULL, -- 계란 특 가격
        big REAL NULL, -- 계란 대 가격
        medium REAL NULL, -- 계란 중 가격
        small REAL NULL -- 계란 소 가격
    );
    -- 오리 일일가격조회
    DROP TABLE IF EXISTS duck_prices;
    CREATE TABLE IF NOT EXISTS duck_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        modymd TEXT NOT NULL, -- 발표일
        typename TEXT NOT NULL, -- 거래종류
        verybig REAL NULL -- 오리(20~26호) 가격
    );
-- 축산물유통정보
    -- 축산물 재고동향
    DROP TABLE IF EXISTS livestock_product_inventory_trends;
    CREATE TABLE IF NOT EXISTS livestock_product_inventory_trends (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        standym TEXT NOT NULL, -- 날짜 구분
        judgekindnm TEXT NOT NULL, -- 축종명
        judgekind TEXT NOT NULL, -- 축종 코드
        livestockpart_1 REAL NULL, -- 소 돼지 안심
        livestockpart_2 REAL NULL, -- 소 돼지 등심
        livestockpart_3 REAL NULL, -- 소 채끝 돼지 전지
        livestockpart_4 REAL NULL, -- 소 목심 돼지 후지
        livestockpart_5 REAL NULL, -- 소 앞다리 돼지 삼겹살
        livestockpart_6 REAL NULL, -- 소 우둔 돼지 목심
        livestockpart_7 REAL NULL, -- 소 설도 돼지 갈비
        livestockpart_8 REAL NULL, -- 소 사태 돼지 특수부위
        livestockpart_9 REAL NULL, -- 소 갈비 돼지 잡육
        livestockpart_10 REAL NULL, -- 소 양지
        livestockpart_11 REAL NULL, -- 소 특수부위
        livestockpart_12 REAL NULL, -- 소 잡육
        totstock REAL NULL, -- 재고량 합계
        unit TEXT NOT NULL, -- 단위
        standyyyy TEXT NOT NULL, -- 기준연도 추가한칼럼
        standmm TEXT NOT NULL -- 기준월 추가한칼럼 
    );
    -- 월별 축산물소비자가격 정보
    DROP TABLE IF EXISTS monthly_livestock_product_prices;
    CREATE TABLE IF NOT EXISTS monthly_livestock_product_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        standym TEXT NOT NULL, -- 기준월
        judgekindnm TEXT NOT NULL, -- 축종명
        itemnm TEXT NOT NULL, -- 품목명
        itemcd TEXT NOT NULL, -- 품목코드
        grdnm TEXT NOT NULL, -- 구분명
        unit TEXT NOT NULL, -- 단위
        ntslprc REAL NULL, -- 가격
        avgyearprc REAL NULL -- 평년가격
    );
    -- 순별 축산물소비자가격 정보
    DROP TABLE IF EXISTS interval_10day_livestock_product_prices;
    CREATE TABLE IF NOT EXISTS interval_10day_livestock_product_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        standym TEXT NOT NULL, -- 기준월
        judgekind TEXT NOT NULL, -- 축종 코드
        judgekindnm TEXT NOT NULL, -- 축종명
        itemnm TEXT NOT NULL, -- 품목명
        itemcd TEXT NOT NULL, -- 품목코드
        grdnm TEXT NOT NULL, -- 구분명
        unit TEXT NOT NULL, -- 단위
        ntslprc_1 REAL NULL, -- 초순가격
        ntslprc_2 REAL NULL, -- 중순가격
        ntslprc_3 REAL NULL, -- 하순가격
        ntslprc REAL NULL, -- 월평균가격
        avgyearprc_1 REAL NULL, -- 평년 초순가격
        avgyearprc_2 REAL NULL, -- 평년 중순가격
        avgyearprc_3 REAL NULL, -- 평년 하순가격
        avgyearprc REAL NULL -- 평년 월평균가격
    );
    -- 연도별 축산물소비자가격 정보
    DROP TABLE IF EXISTS annual_livestock_product_prices;
    CREATE TABLE IF NOT EXISTS annual_livestock_product_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year TEXT NOT NULL, -- 연도
        judgekind TEXT NOT NULL, -- 축종 코드
        judgekindnm TEXT NOT NULL, -- 축종명
        itemnm TEXT NOT NULL, -- 품목명
        itemcd TEXT NOT NULL, -- 품목코드
        grdnm TEXT NOT NULL, -- 구분명
        unit TEXT NOT NULL, -- 단위
        ntslprc REAL NULL, -- 가격
        maxprc REAL NULL, -- 최대가격
        minprc REAL NULL, -- 최소가격
        standy TEXT NOT NULL -- 기준연도 추가한칼럼
    );
    -- 일자별 축산물소비자가격 정보
    DROP TABLE IF EXISTS daily_livestock_product_prices;
    CREATE TABLE IF NOT EXISTS daily_livestock_product_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        standymd TEXT NOT NULL, -- 날짜구분(날짜/평년)
        judgekind TEXT NOT NULL, -- 축종 코드
        judgekindnm TEXT NOT NULL, -- 축종명
        itemnm TEXT NOT NULL, -- 품목명
        itemcd TEXT NOT NULL, -- 품목코드
        grdnm TEXT NOT NULL, -- 구분명
        unit TEXT NOT NULL, -- 단위
        ntslprc REAL NULL, -- 평균가격
        maxprc REAL NULL, -- 최대가격
        minprc REAL NULL, -- 최소가격
        regionprc1 REAL NULL, -- 서울가격
        regionprc2 REAL NULL, -- 부산가격
        regionprc3 REAL NULL, -- 대구가격
        regionprc4 REAL NULL, -- 인천가격
        regionprc5 REAL NULL, -- 광주가격
        regionprc6 REAL NULL, -- 대전가격
        regionprc7 REAL NULL, -- 울산가격
        regionprc8 REAL NULL, -- 세종가격
        regionprc9 REAL NULL, -- 경기가격
        regionprc10 REAL NULL, -- 강원가격
        regionprc11 REAL NULL, -- 충북가격
        regionprc12 REAL NULL, -- 충남가격
        regionprc13 REAL NULL, -- 전북가격
        regionprc14 REAL NULL, -- 전남가격
        regionprc15 REAL NULL, -- 경북가격
        regionprc16 REAL NULL, -- 경남가격
        regionprc17 REAL NULL, -- 제주가격
        standymd_c TEXT NOT NULL -- 기준일자 추가한칼럼
    );
    
    

"""

def init_data_insert():
    for file in Path(os.path.join(root_dir)).rglob("*.xlsx"):
        if '~$' not in file.name and '영문' in file.name :
            sheets = pd.read_excel(file, sheet_name=None)

            for key,df in sheets.items():
                db_table_name = table_config[key]['table_name'] # 국문 시트명으로 DB테이블 명 찾기
                db_table_col = table_config[key]['columns'].keys()

                df = df[db_table_col] # 엑셀에서 읽은 테이블 칼럼 순서 변경

                try:
                    tmp_data = df.itertuples(index=False, name=None)
                    tmp_sql_query = f"INSERT INTO {db_table_name} ({", ".join(db_table_col)}) VALUES ({", ".join(["?"] * len(db_table_col))})"
                    #print(tmp_sql_query)
                    cursor.executemany(
                        tmp_sql_query,
                        tmp_data
                    )
                    print(f'insert 완료 {key} - {db_table_name} ')
                except Exception as e:
                    print(f'에러발생 {e}')


if __name__=='__main__':

    ###############################
    # 초기 DB CREATE + INSERT 설정
    ###############################

    root_dir = r'D:\공공데이터포털\최종결과'

    # DB 연결
    conn = sqlite3.connect(os.path.join(root_dir, "data.db"))
    cursor = conn.cursor()

    # 테이블 생성 명령
    cursor.executescript(CREATE_TABLE_SQL_QUERY)

    # 현재 까지 받은 데이터 DB적재
    init_data_insert()


    # 저장 확정
    conn.commit()

    # 연결 종료
    conn.close()

    # json 저장
    with open(os.path.join(root_dir,"table_config.json"), "w", encoding="utf-8") as f:
        json.dump(table_config, f, ensure_ascii=False, indent=4)