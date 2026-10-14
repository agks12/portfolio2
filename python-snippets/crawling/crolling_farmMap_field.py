from selenium import webdriver
from selenium.webdriver.chrome.service import Service  # Service 임포트
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from tqdm import trange, tqdm
import time
from bs4 import BeautifulSoup
import pandas as pd

# 웹드라이버 경로 설정 (ChromeDriver 예시)
driver = webdriver.Chrome()
driver.get("https://agis.epis.or.kr/ASD/main/intro.do#Wrap")

# 팜맵통계 버튼 누르기 , 버튼을 xpath로 찾음
driver.find_element(By.XPATH,'//*[@id="menu04"]/a').click()
time.sleep(0.5) # 로드 되기 까지 기다리기


# 면적 or 필지수 선택
driver.find_element(By.XPATH,'//*[@id="STA_27"]').click()
time.sleep(0.5) # 로드 되기 까지 기다리기

# 연도는 자동으로 2024 선택 되므로 지역 선택
driver.find_element(By.XPATH,'//*[@id="sido_stat_2025"]').click()
time.sleep(3) # 로드 되기 까지 기다리기

# 시도 선택
select_element = Select(driver.find_element(By.ID, 'sido_stat_2025'))
options = select_element.options

#result = pd.DataFrame(columns = ['시도','시군구','지역','합계(개)','논(개)','밭(개)','과수(개)','시설(개)','인삼(개)'])
result = pd.DataFrame(columns = ['시도','시군구','지역','합계(ha)','논(ha)','밭(ha)','과수(ha)','시설(ha)','인삼(ha)'])
# 시도 각 반복
for option in options:
    value = option.get_attribute('value')

    if value == '00':
        continue  # '선택'은 제외

    select_element.select_by_value(value)
    print(f"{value} - {option.text.strip()} 선택됨")
    # 누르기 해야됨
    # "서울특별시" 선택 (value="11")
    sido_select = Select(driver.find_element(By.ID, 'sido_stat_2025'))
    sido_select.select_by_value(value)

    time.sleep(1)  # AJAX 응답 대기 시간 (조절 가능)
    # 현재 시도 선택됨

    # 시군구 선택
    driver.find_element(By.XPATH, '// *[ @ id = "signgu_stat_2025"]').click()
    time.sleep(0.5)  # 로드 되기 까지 기다리기

    # 시도 선택
    select_element_sgg = Select(driver.find_element(By.ID, 'signgu_stat_2025'))
    options_sgg = select_element_sgg.options

    # 시군구 선택
    for option_sgg in options_sgg:
        value_sgg = option_sgg.get_attribute('value')

        if value_sgg == '00':
            continue  # '선택'은 제외

        select_element_sgg.select_by_value(value_sgg)
        print(f"{value_sgg} - {option_sgg.text.strip()} 선택됨")
        time.sleep(1)  # AJAX 응답 대기 시간 (조절 가능)
        sgg_select = Select(driver.find_element(By.ID, 'signgu_stat_2025'))
        sgg_select.select_by_value(value_sgg)

        # 여기 까지 오면 하나의 시군 하나의 시군구 선택됨

        # 조회하기 누루기
        driver.find_element(By.XPATH, '//*[@id="Wrap"]/div[6]/div[2]/div/div[2]/div[2]/button').click()
        time.sleep(1)  # 로드 되기 까지 기다리기

        table_element = driver.find_element(By.ID, 'farmmapStatTable')
        # 테이블 HTML 가져오기
        html = table_element.get_attribute('outerHTML')

        # BeautifulSoup로 파싱
        soup = BeautifulSoup(html, 'html.parser')

        # 헤더, 데이터 추출
        headers = [th.get_text(strip=True) for th in soup.find_all('th')]

        rows = []
        for tr in soup.find_all('tr')[1:]:
            cells = [td.get_text(strip=True).replace(',', '') for td in tr.find_all('td')]
            rows.append(cells)

        df = pd.DataFrame(rows, columns=headers)

        for col in df.columns[1:]:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        df = df.drop(index=0)

        df['시도'] = option.text.strip()
        df['시군구'] = option_sgg.text.strip()

        result  = pd.concat([result,df],axis=0)


result.to_csv('농경지분류별_농경지_필지_2025_면적.csv',index=False,encoding='utf-8-sig')
