from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from configparser import ConfigParser
from datetime import datetime, timedelta
import os, time, pyperclip, pyautogui, sys
from selenium.webdriver.common.alert import Alert
from datetime import datetime
import re
from webdriver_manager.chrome import ChromeDriverManager
import glob
import pandas as pd

def logWrite(strMessage):
    date = datetime.today().strftime('[%Y/%m/%d %H:%M:%S] ')
    strMessage = date + strMessage
    print(strMessage)
    logFile = open('log.txt', mode='at', encoding='euc-kr')
    logFile.write(strMessage + '\n')
    logFile.close()
    return True

def hasXpath(xpath):
    driver.implicitly_wait(0)
    try:
        driver.find_element_by_xpath(xpath)
        return True
    except:
        return False
    finally:
        driver.implicitly_wait(0.2)

try:
    parser = ConfigParser()
    parser.read('config.ini', encoding='euc-kr')
    work_count = parser.getfloat('대기시간', 'sleep')
    flag = True
except Exception as e:
    strMessage = '[설정파일 Load 실패] = ' + str(e)
    logWrite(strMessage)

try:
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    driver.get('http://eserv-5633.cloudforce.com')
    driver.maximize_window()
except Exception as e:
    strMessage = '[크롬 드라이버 로드 실패] = ' + str(e)
    logWrite(strMessage)

files = glob.glob('.\\input\\*.xlsx')

input_id = driver.find_element(by=By.XPATH, value='//*[@id="username"]')
input_id.send_keys('sungik.jang@ymfk.yokogawa.com')
input_pw = driver.find_element(by=By.XPATH, value='//*[@id="password"]')
input_pw.send_keys('ymfk1234')
ele = driver.find_element(by=By.XPATH, value='//*[@id="Login"]').click()

for file in files:
    df_input = pd.read_excel(file, sheet_name=1)
    for i in df_input.index:

        element = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[1]/section/div[1]/div[1]/one-appnav/div/one-app-nav-bar/nav/div/one-app-nav-bar-item-root[5]/a')))
        driver.execute_script("arguments[0].click();", element)

        time.sleep(work_count)
        element2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="brandBand_1"]/div/div/div/div/div[1]/div[1]/div[2]/ul/li[1]/a'))).click()
        time.sleep(work_count)

        driver.find_element(by=By.XPATH, value='//button[@aria-label="사이트, --없음--"]').click()
        time.sleep(0.5)
        driver.find_element(by=By.XPATH, value=f'//lightning-base-combobox-item[@data-value="{str(df_input[df_input.columns[2]][i])}"]').click()

        driver.find_element(by=By.XPATH, value='//button[@aria-label="대분류, --없음--"]').click()
        time.sleep(0.5)
        driver.find_element(by=By.XPATH, value=f'//lightning-base-combobox-item[@data-value="{str(df_input[df_input.columns[3]][i])}"]').click()

        driver.find_element(by=By.XPATH, value='//button[@aria-label="중분류, --없음--"]').click()
        time.sleep(0.5)
        driver.find_element(by=By.XPATH, value=f'//lightning-base-combobox-item[@data-value="{str(df_input[df_input.columns[4]][i])}"]').click()

        driver.find_element(by=By.XPATH, value='//button[@aria-label="수량 단위, --없음--"]').click()
        time.sleep(0.5)
        driver.find_element(by=By.XPATH, value=f'//lightning-base-combobox-item[@data-value="{str(df_input[df_input.columns[9]][i])}"]').click()

        partName = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Parts_Cd__c"]')
        partName.send_keys(str(df_input[df_input.columns[7]][i]))

        partDetail = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Parts_Detail__c"]')
        partDetail.send_keys(str(df_input[df_input.columns[8]][i]))

        unitPrice = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Unit_Price__c"]')
        unitPrice.send_keys(str(df_input[df_input.columns[10]][i]))

        driver.find_element(by=By.XPATH, value='//button[@aria-label="소모품/정기교환품, --없음--"]').click()
        driver.find_element(by=By.XPATH, value='//button[@aria-label="소모품/정기교환품, --없음--"]').click()
        time.sleep(0.5)
        print(str(df_input[df_input.columns[13]][i]))
        if str(df_input[df_input.columns[13]][i]) == '소모품':
            driver.find_element(by=By.XPATH, value='//lightning-base-combobox-item[@data-value="Expendables"]').click()
        elif str(df_input[df_input.columns[13]][i]) == '순환품':
            driver.find_element(by=By.XPATH, value='//lightning-base-combobox-item[@data-value="Rotation Item"]').click()

        leadTime = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Lead_Time__c"]')
        leadTime.send_keys(re.sub(r'[^0-9]','',str(df_input[df_input.columns[14]][i])))

        orderNum = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Order_Num__c"]')
        orderNum.send_keys(str(df_input[df_input.columns[15]][i]))

        orderPoint = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Order_Point__c"]')
        orderPoint.send_keys(str(df_input[df_input.columns[16]][i]))

        maker = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Maker__c"]')
        maker.send_keys(str(df_input[df_input.columns[17]][i]))

        time.sleep(100)





