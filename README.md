# I. Introduction

회사를 다니는 형이 반복적이고 귀찮은 업무가 생겼는지 나에게 매크로 프로그램을 만들어 줄 수 있는지 일을 맡겨주었다. 어느 코인을 사고 파는 사이트에서 여러 회원의 코인을 자동으로 판매 신청을 해주는 작업이다. 작업의 순서는 파이썬으로 다음과 같이 진행하려고 한다.

1. 회원 정보가 들어있는 excel 파일을 읽는다.
2. 해당 사이트에서 회원 입력 란에 읽었던 파일에서 id 데이터를 입력해준다.
3. 다음 비밀번호 입력 란에 동일한 방법으로 입력해준다.
4. 다음 페이지로 이동하게 되면 여러 메뉴가 있는데 그 중에서 코인거래소 메뉴에 진입한다.
5. 마지막으로 신청수량 란에 회원이 요청했던 판매 코인 수를 excel 파일에서 읽어서 입력한다.
6. 다수의 회원이 있기 때문에 위와 같은 방법으로 하면 크롬 창이 많이 생성되어 메모리가 넘칠 수 있다. 그렇기 때문에 위 과정을 수행하면 크롬 창을 닫아준다.
7. 크롬 창을 닫기 때문에 판매 신청이 정상적으로 이루어졌는지 알기 위해서 엑셀 비고란에 완료 메세지를 입력한다.

# II. Requirement

## II-1. Chromedriver

웹브라우저를 조작하기 위하여 크롬드라이버가 필요하다. **웹드라이버**는 프로그래밍 언어를 이용하여 웹브라우저를 직접적으로 조작할 수 있도록 해 주는 툴입니다.

크롬드라이버를 다운 받기 전에 현재 설치된 크롬의 버전을 알아야 한다.

1. 브라우저 오른쪽 상단의 **점 세개** 클릭
2. **도움말** -> **Chrome 정보**를 선택

![Untitled](https://user-images.githubusercontent.com/56228085/190917022-f3e0188e-bd27-4fd5-8ef3-160ae0cebef0.png)

나는 버전 105.0.5195.125(공식 빌드) (64비트) 를 사용하고 있어

[ChromeDriver - WebDriver for Chrome - Downloads](https://chromedriver.chromium.org/downloads)

해당 사이트에서 105 버전을 다운 받으면 된다.

별도 설치 없이, 해당 파일을 필요한 곳으로 이동하여 사용하면 된다.

![Untitled 1](https://user-images.githubusercontent.com/56228085/190917034-61b7bb93-1056-407c-9b98-cbe329f90f02.png)

## II-2. Selenium

크롬드라이버를 조작하기 위하여 파이썬 Selenium 패키지가 필요하다.

터미널에 해당 코드를 입력하자.

```bash
pip install selenium
```

# III. Login

## III-1. ID & PW 입력하기

입력이나 클릭을 하고 싶은 부분을 특정하기 위하여 검사 도구가 필요하다.

크롬에서 "Ctrl+Shift+C"를 누르면 검사 도구 기능을 이용할 수 있다.

![Untitled 2](https://user-images.githubusercontent.com/56228085/190917044-3ece1239-dfd6-4c12-b08a-ddea5abdad76.png)

혹은 F12를 눌러서 뜨는 개발자 도구 화면에서 왼쪽 상단 커서 모양을 클릭해도 된다.

![Untitled 3](https://user-images.githubusercontent.com/56228085/190917049-e0384d51-9432-43b8-9c55-40cdc6db4018.png)

이제 회원 입력 칸의 코드를 확인한다. 검사 도구를 실행하고, 회원 입력칸을 클릭하면 오른쪽에 해당 소스가 자동으로 표시되는 것을 알 수 있습니다. 여기서는 해당 input field의 id가 `login_id` 라는 점에 주목하자. 전체 코드를 검색해보면 id가 `login_id`인 요소는 회원 입력 칸 밖에 없다.

마찬가지로 password의 id는 `login_pw` 이다.

정리하자면, 입력하고자 하는 id를 찾아 send_keys를 하게 되면 입력을 할 수 있다.

```python
from selenium import webdriver
from selenium.webdriver.common.by import By

driver = webdriver.Chrome("chromedriver")
driver.get('http://office.dhsglobal.biz/User') # 로그인 화면으로 이동
driver.implicitly_wait(15) # 페이지 다 뜰 때 까지 기다림

driver.find_element(By.ID, 'login_id').send_keys('test') # 회원번호
driver.find_element(By.ID, 'login_pw').send_keys("1111") # 비밀번호
```

![Untitled 4](https://user-images.githubusercontent.com/56228085/190917054-53a866ce-48f3-4166-9bf9-b55ad8cf49ce.png)

다음과 같이 selenium module을 불러올 수 없다면 다른 파이썬 버전에 설치되어 있는 경우가 있다.

vscode 오른쪽 아래 파이썬 버전을 클릭하면 인터프리터 선택을 할 수 있는데 다음 추천해주는 파이썬 버전에 selenium module이 설치되어 있어 정상적으로 실행이 가능하다.

![Untitled 5](https://user-images.githubusercontent.com/56228085/190917057-561efe26-b753-4ca3-b325-3c1d5dacb6e5.png)

정상적으로 실행이 되면 다음과 같이 새로운 창이 열리면서 id, pw가 정상적으로 입력이 되었다.

![Untitled 6](https://user-images.githubusercontent.com/56228085/190917169-3e967f05-c0d0-4082-817a-13f4311e1760.png)

## III-2. 로그인 버튼 누르기

검사 도구(Ctrl+Shift+C)를 이용하여 확인 버튼의 소스를 확인한다. 위에서 처럼 고유한 id가 없으므로, XPath를 이용한다. 해당 버튼에 해당하는 소스코드에 마우스 우클릭하여 Copy -> Copy full Xpath를 클릭한다. 

![Untitled 7](https://user-images.githubusercontent.com/56228085/190917066-f4680c14-1955-442d-96c2-835526945fa1.png)

`복사된 XPath:` /html/body/div[3]/div/div/div/div/form/div[4]/button[1]

복사된 XPath를 이용하여 셀레늄에서 해당 버튼을 클릭한다.

위의 코드에 아래의 내용을 추가하면, 회원 ID와 비밀번호를 입력하고 확인 버튼을 누르게 된다.

```python
from selenium import webdriver
from selenium.webdriver.common.by import By

driver = webdriver.Chrome("chromedriver")
driver.get('http://office.dhsglobal.biz/User') # 로그인 화면으로 이동
driver.implicitly_wait(15) # 페이지 다 뜰 때 까지 기다림

driver.find_element(By.ID, 'login_id').send_keys('test') # 회원번호
driver.find_element(By.ID, 'login_pw').send_keys("1111") # 비밀번호

driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div/form/div[4]/button[1]').click()
driver.implicitly_wait(5)
```

# IV. 코인거래소 접속

로그인 후 메뉴 중에 코인거래소 페이지 접속하기 위해서 마찬가지로 버튼만 클릭하면 되기 때문에 코인거래소의 XPath를 복사해준다.

```python
from selenium import webdriver
from selenium.webdriver.common.by import By

driver = webdriver.Chrome("chromedriver")
driver.get('http://office.dhsglobal.biz/User') # 로그인 화면으로 이동
driver.implicitly_wait(15) # 페이지 다 뜰 때 까지 기다림

driver.find_element(By.ID, 'login_id').send_keys('test') # 회원번호
driver.find_element(By.ID, 'login_pw').send_keys("1111") # 비밀번호

driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div/form/div[4]/button[1]').click()
driver.implicitly_wait(5)

driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div[11]/div/div').click()
driver.implicitly_wait(5)
```

# V. 신청 수량 입력

특정 기간에만 신청이 가능하기 때문에 다음과 같은 오류 메세지가 뜬다.

![Untitled 8](https://user-images.githubusercontent.com/56228085/190917072-30b19013-0cb2-42f8-80eb-379031ed5017.png)

그래서인지 신청수량을 입력하는 코드가 동작하지 않는다.

⇒ 신청 기간 때 다시 테스트를 해봐야 할 것 같다.

# VI. Excel 다루기

## VI-1. Excel 읽기

먼저 python을 통해 excel 파일을 읽기 전에 경로를 설정해 주어야 한다.

현재 실행할 .py 경로 디렉터리에 저장해주면 된다. 코인판매 신청.xlsx을 다음 경로에 저장해주었다.

![Untitled 9](https://user-images.githubusercontent.com/56228085/190917076-565f394f-b27d-4d2b-b34a-0974804c9b37.png)

따로 엑셀 파일들이 저장되어 있는 폴더가 있다면 그 폴더의 절대경로를 써주면 된다.

ex) C:\Users\test\Desktop\PythonWorkspace\macro_web\test.xlsx

`openpyxl` 라이브러리를 통해 xlsx 파일을 읽을 수 있다.

terminal에서 openpyxl 라이브러리를 install 하자.

```bash
pip install openpyxl
```

Python에서 pandas로 excel 파일(.xlsx, .xls)을 읽어와 DataFrame에 저장하기 위해서는 `pandas.read_excel()` 함수를 사용하면 된다.

```python
import pandas as pd

# 인덱스를 지정해 시트 설정
user_info = pd.read_excel('코인판매 신청.xlsx', sheet_name=0, engine='openpyxl')

print(user_info)
#     작업날짜        아이디  비밀번호   이름  코인판매 신청수량  비고
# 0    NaN     xxxxxx  1111  송종임        NaN NaN
# 1    NaN     xxxxxx  1111  권옥숙        NaN NaN
# 2    NaN     xxxxxx  1111  노선순        NaN NaN
# 3    NaN     xxxxxx  1111  김보옥        NaN NaN
# 4    NaN     xxxxxx  1111  박주현        NaN NaN
# ..   ...        ...   ...  ...        ...  ..
# 62   NaN     xxxxxx  1111  유춘자        NaN NaN
# 63   NaN     xxxxxx  1111  김춘자        NaN NaN
# 64   NaN     xxxxxx  1111  이귀숙        NaN NaN
# 65   NaN     xxxxxx  1111  이상옥        NaN NaN
# 66   NaN     xxxxxx  1111  장옥녀        NaN NaN
```

xlrd에서 xlsx 확장자를 읽는 기능이 파이썬 3.9 버전 이상부터는 불안정하여 해당 기능 지원이 중단된 것이라 한다. 기본적으로 제공하는 엔진은 xlrd인데 이를  `openpyxl` 엔진으로 변경해주어야 한다.

sheet_name은 index를 넣어도 되고, sheet_name=”시트 이름”과 같이 시트의 이름을 넣어도 된다.

아이디에 해당하는 데이터만 읽고 싶으면 다음 코드를 실행하면 된다.

```python
import pandas as pd

# 인덱스를 지정해 시트 설정
user_info = pd.read_excel('코인판매 신청.xlsx', sheet_name=0, engine='openpyxl')

for i in range(len(user_info)):
    print(user_info["아이디"][i])
```

## VI-2. Excel 쓰기

Excel 파일을 쓰기 위해서 먼저 수정 작업이 이루어질 것이다.

수정은 반복문을 이용해서 해당 인덱스에 접근해 원하는 값을 넣어주면 된다.

```python
import pandas as pd

# 인덱스를 지정해 시트 설정
user_info = pd.read_excel('코인판매 신청.xlsx', sheet_name=0, engine='openpyxl')

for i in range(len(user_info)):
    user_info['비고'][i] = "완료"

# 수정한 data frame(user_info)을 다시 엑셀 파일로 저장
user_info.to_excel('코인판매 신청(완료).xlsx', index=False)
```

`index=False` 를 해주지 않으면 저장된 엑셀 첫번째 열에 index가 저장되기 때문에 False를 해준다.

![Untitled 10](https://user-images.githubusercontent.com/56228085/190917087-23f691e7-8bba-4406-9ccf-1af6081a40bb.png)

위와 같이 새로운 파일이 생성되었다.

# Code

```python
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
```

# Reference

[[Python] 파이썬으로 SRT 예매 프로그램 만들기 (1) 기능 구현하기](https://kminito.tistory.com/79)
