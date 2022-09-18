from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time

# 저장되어 있는 엑셀 파일을 읽어 data frame 구조로 만듦
def read_xlsx(file):
    # 인덱스를 지정해 시트 설정
    return pd.read_excel("{}.xlsx".format(file), sheet_name=0, engine='openpyxl')

# 수정했던 data_frame을 새로운 엑셀 파일로 저장
def write_xlsx(data_frame):
    data_frame.to_excel('코인판매 신청(완료).xlsx', index=False)

# 엑셀 파일에 있는 모든 아이디와 비밀번호 정보를 읽어와 자동으로 판매 신청을 해주는 매크로
def apply_coin_macro(user_info):
    for i in range(len(user_info.head(5))):
        driver = webdriver.Chrome("chromedriver")
        driver.get('http://office.dhsglobal.biz/User') # 로그인 화면으로 이동
        driver.implicitly_wait(5) # 페이지 다 뜰 때 까지 기다림

        driver.find_element(By.ID, 'login_id').send_keys(user_info["아이디"][i]) # 아이디
        driver.find_element(By.ID, 'login_pw').send_keys(int(user_info["비밀번호"][i])) # 비밀번호

        driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div/form/div[4]/button[1]').click() # 버튼 클릭
        driver.implicitly_wait(5)

        driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div[11]/div/div').click()
        driver.implicitly_wait(5)

        # driver.find_element(By.ID, 'req_amt').send_keys("15")
        time.sleep(2)
        driver.quit() # 크롬 창 종료

if __name__=="__main__":
    user_info = read_xlsx("코인판매 신청") # 파일 이름만 (확장자 x)
    apply_coin_macro(user_info)
    # write_xlsx(user_info)