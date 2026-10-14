import sqlite3
import os


# 일자별 종합할 테이블들
daily_table_list = ['pivot_daily_livestock_product_prices',
                    'pivot_pig_representative_price',
                    'pivot_korean_chickens_prices',
                    'pivot_broilers_farm_gate_prices',
                    'pivot_broilers_wholesale_prices',
                    'pivot_egg_prices']

# 주별 종합할 테이블들
weekly_table_list = ['laying_hens_spent_hens_prices']

# 주별 종합할 테이블들
monthly_table_list = ['laying_hens_chicks_prices', 'pivot_livestock_product_inventory_trends', 'pivot_interval_10day_livestock_product_prices']

# 연도별 종합할 테이블들
yearly_table_list = ['pivot_annual_livestock_product_prices']


# --------------------------------------------------
# 테이블 칼럼 정보 반환 -- 피벗 테이블만 형식동일함 피벗테이블 아니면 적용 못할 수도 있음
# --------------------------------------------------
def get_table_columns_info(table_name):
    """
    종합 테이블 만들때 수치 칼럼만 가져오면서 칼럼 이름에 구분할 수 있는 테이블이름 추가하기
    return
    as_columns : select 에 사용될 별칭 칼럼들
    base_ymd_columns : 테이블 기준일자 칼럼
    table_alias : 테이블 별칭
    """
    table_alias = ''.join([s[0] for s in table_name.split('_')]) # 테이블 별칭 - _로 구분한 첫 글자 합침

    cursor.execute(f'PRAGMA table_info("{table_name}")') # sqlite 전용 조회 문법  - id, name, type, notnull, dflt_value, pk 정보 반환
    columns = cursor.fetchall()

    id_idx = [i for i, v in enumerate(columns) if v[1] == 'id']

    base_ymd_idx = 0 if not id_idx else 1 # 첫번째 칼럼이 id인 경우 1번째 칼럼이 날짜, 그 외 0번쨰가 날짜
    base_ymd_columns = columns[base_ymd_idx][1]

    columns_only_real_type = [s[1] for s in columns[base_ymd_idx+1:] if 'ymd' not in s[1]] # type이 real인것만 사용하려했는데 text도 있어서 그냥 전부 사용, 피벗안된것들중 ymd칼럼 제거

    return columns_only_real_type, base_ymd_columns, table_alias


# union 쿼리
def with_all_ymd_query(table_name,base_ymd_columns):
    return f'SELECT {base_ymd_columns} AS ymd FROM {table_name}'

# select 쿼리
def select_alias_columns_query(table_name,columns_only_real_type,table_alias):
    return '\n'.join([f',{table_alias}."{s}" as "{table_name}_{s}" ' for s in columns_only_real_type]) # 칼럼 별칭 적용 (테이블 이름추가된 형태)

# left join 쿼리
def left_join_query(table_name,base_ymd_columns,table_alias):
    return f'LEFT JOIN {table_name} {table_alias}\nON ay.ymd = {table_alias}.{base_ymd_columns}\n' # 무조건 날짜 칼럼 ay.ymd 가정


# --------------------------------------------------
# 각 구문 동적 쿼리 생성
# --------------------------------------------------
def get_dynamic_query(table_list):
    union_query = '' # with query
    select_query = '' # select query
    join_query = '' # left join query

    # 테이블 1개씩 순회하면서 쿼리문 동적으로 구성
    for idx,table_name in enumerate(table_list):
        # 현재 테이블 칼럼 정보 반환
        columns_only_real_type, base_ymd_columns, table_alias = get_table_columns_info(table_name)

        # with 구문에 들어갈 날짜 union 쿼리 구성
        if idx!=0: union_query += '\nUNION\n' # union 구문 추가해주기
        union_query += with_all_ymd_query(table_name,base_ymd_columns)

        # select 구문에 들어갈 별칭 칼럼 쿼리 구성
        select_query += select_alias_columns_query(table_name,columns_only_real_type,table_alias)

        # left join 구문에 들어갈 쿼리 구성
        join_query += left_join_query(table_name,base_ymd_columns,table_alias)

    return union_query, select_query, join_query # select_query 마지막 쉼표 제거


# --------------------------------------------------
# 최종 동적 쿼리 완성본 작성
# --------------------------------------------------
def create_query_script(table,union_query, select_query, join_query):
    return f"""
            DROP TABLE IF EXISTS {table};
            CREATE TABLE IF NOT EXISTS {table} AS
                WITH all_ymd as(
                    {union_query}
                )
                SELECT
                ay.ymd
                {select_query}
                FROM all_ymd ay
                {join_query};
        """


# --------------------------------------------------
# 위 함수들 최상위 함수
# --------------------------------------------------
def common_overall(target_table_list,target_table_name):
    union_query, select_query, join_query = get_dynamic_query(target_table_list)
    return create_query_script(target_table_name,union_query, select_query, join_query)


if __name__=="__main__":
    root_dir = r'D:\공공데이터포털\최종결과'

    # DB 연결
    conn = sqlite3.connect(os.path.join(root_dir, "data.db"))
    cursor = conn.cursor()


    # 테이블 생성 명령
    cursor.executescript((common_overall(daily_table_list,'daily_overall')))
    cursor.executescript((common_overall(weekly_table_list,'weekly_overall')))
    cursor.executescript((common_overall(monthly_table_list,'monthly_overall')))
    cursor.executescript((common_overall(yearly_table_list,'yearly_overall')))


    # 저장 확정
    conn.commit()

    # 연결 종료
    conn.close()