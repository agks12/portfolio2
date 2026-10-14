
import pandas as pd
import sqlite3
import re
import os

root_dir = r'D:\공공데이터포털\최종결과'

# DB 연결
conn = sqlite3.connect(os.path.join(root_dir, "data.db"))

conn.row_factory = sqlite3.Row # pd 할떄 칼럼 보이게하기위한

cursor = conn.cursor()

"""
일자 부분 평년 피벗  : get_table_columns 
join 사용

품종 구분 피벗 : dynamic_pivot
max(case when) 사용 
"""


def dynamic_pivot(
    conn,
    table,
    group_by,
    pivot_columns,
    value_columns,
    base_columns=None,
    where=None,
    order_by=None,
    include_value_name=True,
):
    """
    여러 컬럼의 조합을 기준으로 동적 Pivot

    예:
        pivot_columns = ["축종", "부위"]

        소 + 안심
        소 + 등심
        돼지 + 안심
        돼지 + 등심

        -> 소_안심, 소_등심, 돼지_안심, 돼지_등심

    value_columns가 여러 개면:
        소_안심_price
        소_안심_max
        ...

    value_columns가 1개이고 include_value_name=False면:
        소_안심
        소_등심
        돼지_안심
        돼지_등심
    """

    base_columns = base_columns or []

    # --------------------------------------------------
    # 1. Pivot 대상 컬럼 조합 조회
    # --------------------------------------------------
    pivot_select = ", ".join(
        f'"{column}"'
        for column in pivot_columns
    )

    categories = conn.execute(
        f"""
        SELECT DISTINCT
            {pivot_select}
        FROM "{table}"
        WHERE {" AND ".join(
            f'"{column}" IS NOT NULL'
            for column in pivot_columns
        )}
        ORDER BY {pivot_select}
        """
    ).fetchall()

    if not categories:
        return pd.DataFrame()

    # --------------------------------------------------
    # 2. SELECT 기본 컬럼
    # --------------------------------------------------
    select_columns = [
        f'"{column}"'
        for column in base_columns
    ]

    params = []

    # --------------------------------------------------
    # 3. Pivot 컬럼 생성
    # --------------------------------------------------
    for category in categories:

        # 예:
        # ("소", "안심") -> "소_안심"
        category_name = "_".join(
            str(value)
            for value in category
        )

        for value_column in value_columns:

            if len(value_columns) == 1 and not include_value_name:
                alias = category_name
            else:
                alias = f"{category_name}_{value_column}"

            conditions = " AND ".join(
                f'"{column}" = ?'
                for column in pivot_columns
            )

            select_columns.append(
                f"""
                MAX(
                    CASE
                        WHEN {conditions}
                        THEN "{value_column}"
                    END
                ) AS "{alias}"
                """
            )

            params.extend(category)

    # --------------------------------------------------
    # 4. SQL 생성
    # --------------------------------------------------
    sql = f"""
        SELECT
            {", ".join(select_columns)}
        FROM "{table}"
    """

    if where:
        sql += f"\nWHERE {where}"

    sql += f'\nGROUP BY "{group_by}"'

    if order_by:
        sql += f"\nORDER BY {order_by}"
    else:
        sql += f'\nORDER BY "{group_by}"'

    # --------------------------------------------------
    # 5. 실행 -> DataFrame
    # --------------------------------------------------
    result = conn.execute(sql, params).fetchall()

    return pd.DataFrame([dict(row)for row in result])


def get_table_columns(conn, table, prefix=None):
    """
    평년 날짜 피벗할 때 필요한 칼럼 정보 반환
    """
    columns = conn.execute(
        f'PRAGMA table_info("{table}")'
    ).fetchall()

    result = [row[1] for row in columns]

    if prefix:
        result = [
            column
            for column in result
            if column.startswith(prefix)
        ]

    return result


def save_dataframe_to_db(df, conn, table_name):
    """
    SELECT로 얻은 피벗 판다스 테이블 DB에 저장
    """
    df.to_sql(
        table_name,
        conn,
        if_exists="replace",
        index=False
    )

    print(f'{table_name} 저장')


def get_case_when_pivot(conn,table:str,base_columns:str,pivot_columns:list[str],value_columns:list[str]):
    """
    품종 구분 피벗하는 case when 방식은 동일하기 때문에 함수로 적용 가능
    pivot_columns : 피벗에 사용되는 기준 칼럼(품종,축종,등급..)
    base_columns : 보여줄 즉 남길 칼럼
    """
    return dynamic_pivot(
        conn=conn,
        table=table,
        group_by=base_columns,
        pivot_columns=pivot_columns,
        value_columns=value_columns,
        base_columns=[base_columns],
    )


def get_daily_livestock_product_prices(conn):
    all_columns = get_table_columns(
        conn,
        "daily_livestock_product_prices"
    )

    price_columns = [
        column
        for column in all_columns
        if column in {
            "ntslprc",
            "maxprc",
            "minprc",
        }
    ]

    if not price_columns:
        return []

    avg_year_current_columns = ",\n".join(
        f'a."{column}"'
        for column in price_columns
    )

    avg_year_average_columns = ",\n".join(
        f'b."{column}" AS "avgyearprc_{column}"'
        for column in price_columns
    )

    region_columns = get_table_columns(
        conn,
        "daily_livestock_product_prices",
        prefix="regionprc"
    )

    region_columns = sorted(
        [
            column
            for column in region_columns
            if column[len("regionprc"):].isdigit()
        ],
        key=lambda x: int(x.replace("regionprc", ""))
    )

    if not region_columns:
        return []

    current_columns = ",\n".join(
        f'a."{column}"'
        for column in region_columns
    )

    average_columns = ",\n".join(
        f'b."{column}" AS "avgyear_{column}"'
        for column in region_columns
    )

    sql = f"""
        SELECT
            a.standymd,
            a.judgekind,
            a.judgekindnm,
            a.itemnm,
            a.itemcd,
            a.grdnm,
            a.unit,
            a.ntslprc,
            a.maxprc,
            a.minprc,
            
            {avg_year_current_columns},
            {current_columns},

            {avg_year_average_columns},
            {average_columns}

        FROM daily_livestock_product_prices a

        LEFT JOIN daily_livestock_product_prices b
            ON a.judgekind = b.judgekind
            AND a.itemcd = b.itemcd
            AND a.grdnm = b.grdnm
            AND a.standymd_c = b.standymd_c
            AND b.standymd = '평년'

        WHERE a.standymd != '평년'

        ORDER BY a.standymd
    """

    # 피벗으로 추가된 값 칼럼
    prc_col = [avg_year_current_columns,current_columns,avg_year_average_columns,average_columns]

    return_col = []
    for item in prc_col:
        matches = re.findall(r'"([^"]+)"', item)
        return_col.extend(matches)

    return_col = list(dict.fromkeys(return_col))

    result = conn.execute(sql).fetchall()
    return pd.DataFrame([dict(row) for row in result]) , return_col


def get_annual_livestock_product_prices(conn):
    all_columns = get_table_columns(
        conn,
        "annual_livestock_product_prices"
    )

    price_columns = [
        column
        for column in all_columns
        if column in {
            "ntslprc",
            "maxprc",
            "minprc",
        }
    ]

    if not price_columns:
        return []

    current_columns = ",\n".join(
        f'a."{column}"'
        for column in price_columns
    )

    average_columns = ",\n".join(
        f'b."{column}" AS "avgyearprc_{column}"'
        for column in price_columns
    )

    sql = f"""
        SELECT
            a.year,
            a.judgekind,
            a.judgekindnm,
            a.itemnm,
            a.itemcd,
            a.grdnm,
            a.unit,

            {current_columns},

            {average_columns}

        FROM annual_livestock_product_prices a

        LEFT JOIN annual_livestock_product_prices b
            ON a.judgekind = b.judgekind
            AND a.itemcd = b.itemcd
            AND a.grdnm = b.grdnm
            AND a.standy = b.standy
            AND b.year = '평년'

        WHERE a.year != '평년'

        ORDER BY a.year
    """

    # 피벗으로 추가된 값 칼럼
    prc_col = [current_columns,average_columns]

    return_col = []
    for item in prc_col:
        matches = re.findall(r'"([^"]+)"', item)
        return_col.extend(matches)

    return_col = list(dict.fromkeys(return_col))

    result = conn.execute(sql).fetchall()
    return pd.DataFrame([dict(row) for row in result]), return_col

# 형식 변경 필요한 테이블 일일이 변경
def directly_process_pivot_livestock_product_inventory_trends():
    f = """
        UPDATE pivot_livestock_product_inventory_trends SET standym = SUBSTR(standym, 1, 4) || SUBSTR(standym, 6, 2)
    """
    conn.execute(f)
    print(f'pivot_livestock_product_inventory_trends standym 칼럼 변경')


if __name__=="__main__":

    # SELECT 피벗 테이블
    df1 = get_case_when_pivot(conn, "pig_representative_price", "sumymd", ["sablegubn"], ["costamt"])# 돈육 대표가격현황
    df2 = get_case_when_pivot(conn, "korean_chickens_prices", "modymd", ["typename"], ["agency", "mart", "franchise", "average"]) # 토종닭 가격조회
    df3 = get_case_when_pivot(conn, "broilers_farm_gate_prices", "modymd", ["typename"], ["big", "medium", "small"])# 육계(산지) 일일가격조회
    df4 = get_case_when_pivot(conn, "broilers_wholesale_prices", "modymd", ["typename"], ["franchise","mart","agency","avg"]) # 육계(산지) 일일가격조회
    df5 = get_case_when_pivot(conn, "egg_prices", "modymd", ["typename"], ["verybig", "special", "big", "medium", "small"]) # 계란 일일가격조회
    df6 = get_case_when_pivot(conn, "interval_10day_livestock_product_prices", "standym", ["judgekindnm","itemnm","grdnm"], ["ntslprc_1","ntslprc_2","ntslprc_3","ntslprc","avgyearprc_1","avgyearprc_2","avgyearprc_3","avgyearprc"]) # 순별 축산물소비자가격 정보
    df7 = get_case_when_pivot(conn, "livestock_product_inventory_trends", "standym", ["judgekindnm","unit"], ["livestockpart_1", "livestockpart_2", "livestockpart_3", "livestockpart_4", "livestockpart_5", "livestockpart_6", "livestockpart_7", "livestockpart_8", "livestockpart_9", "livestockpart_10", "livestockpart_11", "livestockpart_12", "totstock"]) # 축산물 재고동향
    df8, col_daily_livestock = get_daily_livestock_product_prices(conn) # 일자별 축산물소비자가격 정보 - 평년 피벗
    df9, col_annual_livestock = get_annual_livestock_product_prices(conn) # 연도별 축산물소비자가격 정보 - 평년 피벗
    #
    # # DB 저장
    save_dataframe_to_db(df1,conn,'pivot_pig_representative_price')
    save_dataframe_to_db(df2,conn,'pivot_korean_chickens_prices')
    save_dataframe_to_db(df3,conn,'pivot_broilers_farm_gate_prices')
    save_dataframe_to_db(df4,conn,'pivot_broilers_wholesale_prices')
    save_dataframe_to_db(df5,conn,'pivot_egg_prices')
    save_dataframe_to_db(df6,conn,'pivot_interval_10day_livestock_product_prices')
    save_dataframe_to_db(df7,conn,'pivot_livestock_product_inventory_trends')
    save_dataframe_to_db(df8,conn,'pivot_daily_livestock_product_prices')
    save_dataframe_to_db(df9,conn,'pivot_annual_livestock_product_prices')
    #
    # # 평년 피벗 중 칼럼 피벗 재적용
    df8 = get_case_when_pivot(conn,"pivot_daily_livestock_product_prices","standymd", ["judgekindnm","itemnm","grdnm"], col_daily_livestock) # 일자별 축산물소비자가격 정보 - 품종으로 피벗
    df9 = get_case_when_pivot(conn,"pivot_annual_livestock_product_prices","year", ["judgekindnm","itemnm","grdnm"], col_annual_livestock) # 연도별 축산물소비자가격 정보 - 품종으로 피벗
    save_dataframe_to_db(df8,conn,'pivot_daily_livestock_product_prices')
    save_dataframe_to_db(df9,conn,'pivot_annual_livestock_product_prices')
    #
    #
    # # 날짜 형식 동일 하게 처리 -- 별로 없어서 이 파일에서 적용 pivot_livestock_product_inventory_trends 테이블 standym 칼럼 yyyy년mm월 -> yyyymm으로 변경
    directly_process_pivot_livestock_product_inventory_trends()

    # 저장 확정 -- 없어도 저장됨
    conn.commit()
    # 연결 종료
    conn.close()