import imp
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import datetime as dt
import time, sys
import re
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import Qt, QCoreApplication
from PyQt5.QtGui import QDoubleValidator, QStandardItemModel, QIcon, QStandardItem, QIntValidator, QFont
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QProgressBar, QPlainTextEdit, QWidget, QGridLayout, QGroupBox, QLineEdit, QSizePolicy, QToolButton, QLabel, QFrame, QListView, QMenuBar, QStatusBar, QPushButton, QApplication, QCalendarWidget, QVBoxLayout, QFileDialog, QCheckBox
from PyQt5.QtCore import pyqtSlot, pyqtSignal, QObject, QThread, QRect, QSize, QDate
import logging
from logging.handlers import RotatingFileHandler

class ThreadClass(QObject):
    returnTotalCnt = pyqtSignal(int)
    returnI = pyqtSignal(int)
    returnError = pyqtSignal(Exception)
    returnEnd = pyqtSignal(bool)
    returnMessage = pyqtSignal(str)

    def __init__(self, username, pw, filepath):
        super().__init__()
        self.isRunning = True
        self.username = username
        self.pw = pw
        self.filepath = filepath

    def hasXpath(self, driver, xpath):
        driver.implicitly_wait(0)
        try:
            driver.find_element(by=By.XPATH, value=xpath)
            return True
        except:
            return False
        finally:
            driver.implicitly_wait(0.2)  
    
    def inputPartData(self, driver, df_input, i):
        try:
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="brandBand_1"]/div/div/div/div/div[1]/div[1]/div[2]/ul/li[1]/a'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//button[@aria-label="사이트, --없음--"]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, f'//lightning-base-combobox-item[@data-value="{str(df_input[df_input.columns[2]][i])}"]'))).click()

            driver.find_element(by=By.XPATH, value='//button[@aria-label="대분류, --없음--"]').click()
            time.sleep(0.5)
            driver.find_element(by=By.XPATH, value=f'//lightning-base-combobox-item[@data-value="{str(df_input[df_input.columns[3]][i])}"]').click()

            if str(df_input[df_input.columns[4]][i]) != '' and str(df_input[df_input.columns[4]][i]) != 'nan':
                driver.find_element(by=By.XPATH, value='//button[@aria-label="중분류, --없음--"]').click()
                time.sleep(0.5)
                driver.find_element(by=By.XPATH, value=f'//lightning-base-combobox-item[@data-value="{str(df_input[df_input.columns[4]][i])}"]').click()

            if str(df_input[df_input.columns[9]][i]) != '' and str(df_input[df_input.columns[9]][i]) != 'nan':
                driver.find_element(by=By.XPATH, value='//button[@aria-label="수량 단위, --없음--"]').click()
                time.sleep(0.5)
                driver.find_element(by=By.XPATH, value=f'//lightning-base-combobox-item[@data-value="{str(df_input[df_input.columns[9]][i])}"]').click()

            if str(df_input[df_input.columns[15]][i]) != '' and str(df_input[df_input.columns[15]][i]) != 'nan':
                leadTime = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Lead_Time__c"]')
                leadTime.send_keys(re.sub(r'[^0-9]','',str(int(df_input[df_input.columns[15]][i]))))
                time.sleep(0.5)
            if str(df_input[df_input.columns[16]][i]) != '' and str(df_input[df_input.columns[16]][i]) != 'nan':
                orderNum = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Order_Num__c"]')
                orderNum.send_keys(re.sub(r'[^0-9]','',str(int(df_input[df_input.columns[16]][i]))))
                time.sleep(0.5)
            if str(df_input[df_input.columns[17]][i]) != '' and str(df_input[df_input.columns[17]][i]) != 'nan':
                orderPoint = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Order_Point__c"]')
                orderPoint.send_keys(re.sub(r'[^0-9]','',str(int(df_input[df_input.columns[17]][i]))))
                time.sleep(0.5)


            if str(df_input[df_input.columns[6]][i]) != '' and str(df_input[df_input.columns[6]][i]) != 'nan':
                partCd = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Parts_Cd__c"]')
                partCd.send_keys(df_input[df_input.columns[6]][i])
                time.sleep(0.5)

            if str(df_input[df_input.columns[7]][i]) != '' and str(df_input[df_input.columns[7]][i]) != 'nan':
                partName = driver.find_element(by=By.XPATH, value='//input[@name="Name"]')
                partName.send_keys(str(df_input[df_input.columns[7]][i]))
                time.sleep(0.5)
            
            if str(df_input[df_input.columns[8]][i]) != '' and str(df_input[df_input.columns[8]][i]) != 'nan':
                partDetail = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Parts_Detail__c"]')
                partDetail.send_keys(str(df_input[df_input.columns[8]][i]))
                time.sleep(0.5)
            if str(df_input[df_input.columns[18]][i]) != '' and str(df_input[df_input.columns[18]][i]) != 'nan':
                maker = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Maker__c"]')
                maker.send_keys(str(df_input[df_input.columns[18]][i]))
                time.sleep(0.5)
            if str(df_input[df_input.columns[19]][i]) != '' and str(df_input[df_input.columns[19]][i]) != 'nan':
                makerModel = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Maker_Model__c"]')
                makerModel.send_keys(str(df_input[df_input.columns[19]][i]))
                time.sleep(0.5)

            admin = driver.find_element(by=By.XPATH, value='/html/body/div[4]/div[2]/div/div[2]/div/div[2]/div/div/div/records-modal-lwc-detail-panel-wrapper/records-record-layout-event-broker/slot/records-lwc-detail-panel/records-base-record-form/div/div/div/div/records-lwc-record-layout/forcegenerated-detailpanel_eserv__parts__c___012000000000000aaa___full___create___recordlayout2/records-record-layout-block/slot/records-record-layout-section[1]/div/div/div/slot/records-record-layout-row[9]/slot/records-record-layout-item[2]/div/span/slot/records-record-layout-lookup/lightning-lookup/lightning-lookup-desktop/lightning-grouped-combobox/div/div/lightning-base-combobox/div/div[1]/input')
            admin.send_keys('장 성익')
            time.sleep(1)
            admin.send_keys(Keys.ENTER)
            time.sleep(1)
            driver.find_element(by=By.XPATH, value='//a[text()="장 성익"]').click()
            time.sleep(1)

            list_currency = ['없음','JPY','KRW','USD']
            for currency in list_currency:
                if currency in str(df_input[df_input.columns[11]][i]) :
                    if currency != '없음':
                        driver.find_element(by=By.XPATH, value='//button[@aria-label="통화, KRW - 대한민국 원"]').click()
                        time.sleep(0.5)
                        driver.find_element(by=By.XPATH, value=f'//lightning-base-combobox-item[@data-value="{currency}"]').click()
                        time.sleep(0.5)
            driver.execute_script("arguments[0].scrollIntoView();", admin)
            if str(df_input[df_input.columns[14]][i]) == '소모품':
                driver.find_element(by=By.XPATH, value='//button[@aria-label="소모품/정기교환품, --없음--"]').click()
                time.sleep(0.5)
                driver.find_element(by=By.XPATH, value='//lightning-base-combobox-item[@data-value="Expendables"]').click()
                time.sleep(0.5)
            elif str(df_input[df_input.columns[14]][i]) == '순환품':
                driver.find_element(by=By.XPATH, value='//button[@aria-label="소모품/정기교환품, --없음--"]').click()
                time.sleep(0.5)
                driver.find_element(by=By.XPATH, value='//lightning-base-combobox-item[@data-value="Rotation Item"]').click()
                time.sleep(0.5)

            if str(df_input[df_input.columns[10]][i]) != '' and str(df_input[df_input.columns[10]][i]) != 'nan':
                unitPrice = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Unit_Price__c"]')
                unitPrice.send_keys(re.sub(r'[^0-9]','', str(int(df_input[df_input.columns[10]][i]))))
                time.sleep(0.5)

            if str(df_input[df_input.columns[12]][i]) == 'X' :
                ele = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Regular__c"]')
                driver.execute_script("arguments[0].click();", ele)
                time.sleep(0.5)
            if str(df_input[df_input.columns[13]][i]) == 'O' :
                ele = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Unused__c"]')
                driver.execute_script("arguments[0].click();", ele)
                time.sleep(0.5)

            if str(df_input[df_input.columns[20]][i]).replace('-','') != '' and str(df_input[df_input.columns[20]][i]) != 'NaT' and str(df_input[df_input.columns[20]][i]) != 'nan':
                stopDt = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Stop_Dt__c"]')
                stopDt.send_keys(df_input[df_input.columns[20]][i].strftime('%Y. %m. %d.'))
                time.sleep(0.5)
            if str(df_input[df_input.columns[21]][i]) != '' and str(df_input[df_input.columns[21]][i]) != 'nan':
                supplier = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Supplier__c"]')
                supplier.send_keys(str(df_input[df_input.columns[21]][i]))             
                time.sleep(0.5)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//button[@name="SaveEdit"]'))).click()
            time.sleep(1)

        except Exception as e:
            self.returnError.emit(e)
            self.thread().quit()
            return
    
    def run(self):
        try:
            options = webdriver.ChromeOptions()
            options.add_experimental_option("excludeSwitches", ["enable-logging"])
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
            driver.get('http://eserv-5633.cloudforce.com')
            driver.maximize_window()

            input_id = driver.find_element(by=By.XPATH, value='//*[@id="username"]')
            input_id.send_keys(self.username)
            input_pw = driver.find_element(by=By.XPATH, value='//*[@id="password"]')
            input_pw.send_keys(self.pw)
            driver.find_element(by=By.XPATH, value='//*[@id="Login"]').click()

            df_input = pd.read_excel(self.filepath , sheet_name='N-1', skiprows=7)
            df_output = pd.DataFrame(columns=df_input.columns)
            today = datetime.today().strftime('%Y%m%d%H%M%S')
            totalCnt = df_input.index[-1] * 2
            self.returnTotalCnt.emit(totalCnt)
            for i in df_input.index:
                element = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[1]/section/div[1]/div[1]/one-appnav/div/one-app-nav-bar/nav/div/one-app-nav-bar-item-root[5]/a')))
                driver.execute_script("arguments[0].click();", element)
                searchInput = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.XPATH, '//input[@name="eserv__Parts__c-search-input"]')))
                searchInput.clear()
                time.sleep(0.5)
                searchInput.send_keys(str(df_input[df_input.columns[7]][i]))
                time.sleep(0.5)
                searchInput.send_keys(Keys.ENTER)
                time.sleep(2)
                for j in range(1,999):
                    xpathDefault = '//*[@id="brandBand_1"]/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div/table/tbody/tr/th/span/a'
                    xpath = f'//*[@id="brandBand_1"]/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div/table/tbody/tr[{str(j)}]/th/span/a'
                    if self.hasXpath(driver, xpath):
                        if str(df_input[df_input.columns[7]][i]).strip() == driver.find_element(by=By.XPATH, value=xpath).text:
                            if str(df_input[df_input.columns[3]][i]).strip() == driver.find_element(by=By.XPATH, value=f'/html/body/div[4]/div[1]/section/div[1]/div[2]/div[1]/div/div/div/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div/table/tbody/tr[{str(j)}]/td[4]/span/span[1]').text:
                                if str(df_input[df_input.columns[4]][i]).strip().replace('nan','') == driver.find_element(by=By.XPATH, value=f'/html/body/div[4]/div[1]/section/div[1]/div[2]/div[1]/div/div/div/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div/table/tbody/tr[{str(j)}]/td[5]/span/span[1]').text:
                                    if str(df_input[df_input.columns[8]][i]).strip().replace('nan','') == driver.find_element(by=By.XPATH, value=f'/html/body/div[4]/div[1]/section/div[1]/div[2]/div[1]/div/div/div/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div/table/tbody/tr[{str(j)}]/td[7]/span/span[1]').text:
                                        break
                    # elif j == 1 and self.hasXpath(driver, xpathDefault):
                    #     if str(df_input[df_input.columns[7]][i]).strip() != driver.find_element(by=By.XPATH, value=xpathDefault).text:
                    #         if str(df_input[df_input.columns[3]][i]).strip() != driver.find_element(by=By.XPATH, value=f'/html/body/div[4]/div[1]/section/div[1]/div[2]/div[1]/div/div/div/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div/table/tbody/tr/td[4]/span/span[1]').text:
                    #             if str(df_input[df_input.columns[4]][i]).strip().replace('nan','') != driver.find_element(by=By.XPATH, value=f'/html/body/div[4]/div[1]/section/div[1]/div[2]/div[1]/div/div/div/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div/table/tbody/tr/td[5]/span/span[1]').text:
                    #                 if str(df_input[df_input.columns[8]][i]).strip().replace('nan','') != driver.find_element(by=By.XPATH, value=f'/html/body/div[4]/div[1]/section/div[1]/div[2]/div[1]/div/div/div/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div/table/tbody/tr/td[7]/span/span[1]').text:
                    #                     self.inputPartData(driver, df_input, i)
                    #                     break
                    #                 else:
                    #                     break
                    #             else:
                    #                 break
                    #         else:
                    #             break
                    #     else:
                    #         break
                    else:
                        self.inputPartData(driver, df_input, i)
                        break
                self.returnI.emit(i)

            for i in df_input.index:
                element = WebDriverWait(driver, 500).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div[1]/section/div[1]/div[1]/one-appnav/div/one-app-nav-bar/nav/div/one-app-nav-bar-item-root[6]/a')))
                driver.execute_script("arguments[0].click();", element)
                time.sleep(1.5)

                for j in range(1,999):
                    # xpath_partName = f'//*[@id="brandBand_1"]/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div/table/tbody/tr[{str(j)}]/td[9]/span/a'
                    # xpath_partDetail = f'/html/body/div[4]/div[1]/section/div[1]/div[2]/div[1]/div/div/div/div/div/div/div/div[2]/div/div[1]/div[2]/div[2]/div[1]/div/div/table/tbody/tr[{str(j)}]/td[10]/span/span[1]'
                    # if self.hasXpath(driver, xpath_partName):
                    #     ele_partName = driver.find_element(by=By.XPATH, value=xpath_partName)
                    #     text = ele_partName.text
                    #     partDetail = str(df_input[df_input.columns[8]][i]).strip().replace('nan','').replace('\n','')
                    #     if self.hasXpath(driver, xpath_partDetail):
                    #         text_partDetail = driver.find_element(by=By.XPATH, value=xpath_partDetail).text
                    #     else:
                    #         text_partDetail = ''
                    #     ele_partName.send_keys(Keys.DOWN)
                    #     if text == str(df_input[df_input.columns[7]][i]).strip() and text_partDetail == partDetail:
                    #         self.returnMessage.emit(f'「{str(df_input[df_input.columns[7]][i])}」부품이 이미 등록되어 있습니다. 재고미등록 리스트는 Output폴더를 확인해주세요.')
                    #         df_output = df_output.append(df_input.iloc[i])   
                    #         df_output = df_output.reset_index(drop=True)
                    #         df_output.to_excel('.\\Output\\재고미등록리스트_'+str(today)+'.xlsx')
                    #         break
                    # else:
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="brandBand_1"]/div/div/div/div/div[1]/div[1]/div[2]/ul/li[1]/a'))).click()
                        time.sleep(1)
                        if str(df_input[df_input.columns[29]][i]) != '' and str(df_input[df_input.columns[29]][i]) != 'nan':
                                driver.find_element(by=By.XPATH, value='//input[@name="eserv__Order_Point__c"]').send_keys(str(df_input[df_input.columns[26]][i]).replace('-',''))
                                time.sleep(0.5)
                        driver.find_element(by=By.XPATH, value='//button[@aria-label="사이트, --없음--"]').click()
                        time.sleep(0.5)
                        driver.find_element(by=By.XPATH, value=f'//lightning-base-combobox-item[@data-value="{str(df_input[df_input.columns[2]][i])}"]').click()
                        driver.find_element(by=By.XPATH, value='//button[@aria-label="창고, --없음--"]').click()
                        time.sleep(0.5)
                        driver.find_element(by=By.XPATH, value=f'//lightning-base-combobox-item[@data-value="{str(df_input[df_input.columns[22]][i])}"]').click()
                        time.sleep(0.5)
                        if str(df_input[df_input.columns[2]][i]) != '' and str(df_input[df_input.columns[2]][i]) != 'nan':
                            driver.find_element(by=By.XPATH, value='//input[@name="eserv__Shelf__c"]').send_keys(df_input[df_input.columns[1]][i])
                            time.sleep(0.5)
                        
                        if str(df_input[df_input.columns[23]][i]) != '' and str(df_input[df_input.columns[23]][i]) != 'nan':
                            qty = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Qty__c"]')
                            qty.clear()
                            time.sleep(0.5)
                            qty.send_keys(re.sub(r'[^0-9]','', str(int(df_input[df_input.columns[23]][i]))))
                        
                        if str(df_input[df_input.columns[24]][i]) != '' and str(df_input[df_input.columns[24]][i]) != 'nan':
                            qty = driver.find_element(by=By.XPATH, value='//input[@name="eserv__Unit_Price__c"]')
                            qty.clear()
                            time.sleep(0.5)
                            qty.send_keys(re.sub(r'[^0-9]','', str(int(df_input[df_input.columns[24]][i]))))

                        list_currency = ['없음','JPY','KRW','USD']
                        for currency in list_currency:
                            if currency in str(df_input[df_input.columns[11]][i]) :
                                if currency != '없음':
                                    driver.find_element(by=By.XPATH, value='//button[@aria-label="통화, KRW - 대한민국 원"]').click()
                                    driver.find_element(by=By.XPATH, value='//button[@aria-label="통화, KRW - 대한민국 원"]').click()
                                    time.sleep(0.5)
                                    driver.find_element(by=By.XPATH, value=f'//lightning-base-combobox-item[@data-value="{currency}"]').click()
                                    time.sleep(0.5)

                        master = driver.find_element(by=By.XPATH, value='/html/body/div[4]/div[2]/div/div[2]/div/div[2]/div/div/div/records-modal-lwc-detail-panel-wrapper/records-record-layout-event-broker/slot/records-lwc-detail-panel/records-base-record-form/div/div/div/div/records-lwc-record-layout/forcegenerated-detailpanel_eserv__place__c___012000000000000aaa___full___create___recordlayout2/records-record-layout-block/slot/records-record-layout-section/div/div/div/slot/records-record-layout-row[1]/slot/records-record-layout-item[2]/div/span/slot/records-record-layout-lookup/lightning-lookup/lightning-lookup-desktop/lightning-grouped-combobox/div[1]/div/lightning-base-combobox/div/div[1]/input')
                        # if len(str(df_input[df_input.columns[8]][i]).strip().replace('nan','').replace('\n',''))>0:
                        #     master.send_keys(str(df_input[df_input.columns[8]][i]).strip().replace('nan','').replace('\n',''))
                        # else:
                        master.send_keys(str(df_input[df_input.columns[7]][i]).strip())
                        time.sleep(1)
                        master.send_keys(Keys.ENTER)
                        time.sleep(2)
                        master.send_keys(Keys.ENTER)
                        
                        time.sleep(2)

                        # for j in range(1,999):
                        #     xpath_partName = f'/html/body/div[4]/div[2]/div[5]/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/table/tbody/tr[{str(j)}]/td[1]/a'
                        #     if self.hasXpath(driver, xpath_partName):
                        #         ele_partName = driver.find_element(by=By.XPATH, value=xpath_partName)
                        #         text_partname = ele_partName.text
                        #         partDetail = str(df_input[df_input.columns[8]][i]).strip().replace('nan','').replace('\n','')
                        #         if len(partDetail) > 0:
                        #             xpath_partDetail = f'/html/body/div[4]/div[2]/div[5]/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[2]/div/div/table/tbody/tr[{str(j)}]/td[6]/span'
                        #             if self.hasXpath(driver, xpath_partDetail):
                        #                 text_partDetail = driver.find_element(by=By.XPATH, value=xpath_partDetail).text
                        #                 if partDetail == text_partDetail:
                        #                     ele_partName.click()
                        #         else:
                        #             if text_partname == str(df_input[df_input.columns[7]][i]):
                        #                 ele_partName.click()
                        #     else:
                        #         driver.find_element(by=By.XPATH, value='/html/body/div[4]/div[2]/div[2]/div[2]/div/div[1]/button').click()
                        #         self.returnMessage.emit(f'{str(df_input[df_input.columns[7]][i]).strip()} 가 부품마스터에 등록되지 않았습니다. 확인해주세요')        
                        #         break
                        xpath1 = f'//a[text()="{str(df_input[df_input.columns[7]][i]).strip()}"]'
                        xpath2 = f'//span[text()="{str(df_input[df_input.columns[3]][i]).strip()}"]'

                        # category = str(df_input[df_input.columns[4]][i]).strip().replace('nan','')
                        partDetail = str(df_input[df_input.columns[8]][i]).strip().replace('nan','').replace('\n','')
                        if len(partDetail) > 0:
                            xpath4 = f'//span[text()="{partDetail}"]'
                            if self.hasXpath(driver, xpath1) and self.hasXpath(driver, xpath2) and self.hasXpath(driver, xpath4):
                                driver.execute_script("arguments[0].click();", driver.find_element(by=By.XPATH, value=xpath1))
                            else:
                                driver.find_element(by=By.XPATH, value='/html/body/div[4]/div[2]/div[2]/div[2]/div/div[1]/button').click()
                                self.returnMessage.emit(f'「{str(df_input[df_input.columns[7]][i]).strip()}」 가 부품마스터에 등록되지 않았습니다. 확인해주세요') 
                        else:
                            if self.hasXpath(driver, xpath1) and self.hasXpath(driver, xpath2):
                                driver.execute_script("arguments[0].click();", driver.find_element(by=By.XPATH, value=xpath1))
                            else:
                                driver.find_element(by=By.XPATH, value='/html/body/div[4]/div[2]/div[2]/div[2]/div/div[1]/button').click()
                                self.returnMessage.emit(f'「{str(df_input[df_input.columns[7]][i]).strip()}」 가 부품마스터에 등록되지 않았습니다. 확인해주세요') 
                        
                        # if self.hasXpath(driver, xpath1) and self.hasXpath(driver, xpath2):
                        #     if len(category) > 0:
                        #         xpath3 = f'//span[text()="{category}"]'
                        #         if self.hasXpath(driver, xpath3):
                        #             if len(partDetail) > 0: 
                        #                 xpath4 = f'//span[text()="{partDetail}"]'
                        #                 if self.hasXpath(driver, xpath4):
                        #                     driver.find_element(by=By.XPATH, value=xpath1).click()
                        #                 else:
                        #                     driver.find_element(by=By.XPATH, value='/html/body/div[4]/div[2]/div[2]/div[2]/div/div[1]/button').click()
                        #                     self.returnMessage.emit(f'{str(df_input[df_input.columns[7]][i]).strip()} 가 부품마스터에 등록되지 않았습니다. 확인해주세요')
                        #             else:
                        #                 driver.find_element(by=By.XPATH, value=xpath1).click()
                        #         else:
                        #             driver.find_element(by=By.XPATH, value='/html/body/div[4]/div[2]/div[2]/div[2]/div/div[1]/button').click()
                        #             self.returnMessage.emit(f'{str(df_input[df_input.columns[7]][i]).strip()} 가 부품마스터에 등록되지 않았습니다. 확인해주세요') 
                        #     elif len(partDetail) > 0:
                        #         xpath4 = f'//span[text()="{partDetail}"]'
                        #         if self.hasXpath(driver, xpath4):
                        #             driver.find_element(by=By.XPATH, value=xpath1).click()
                        #         else:
                        #             driver.find_element(by=By.XPATH, value='/html/body/div[4]/div[2]/div[2]/div[2]/div/div[1]/button').click()
                        #             self.returnMessage.emit(f'{str(df_input[df_input.columns[7]][i]).strip()} 가 부품마스터에 등록되지 않았습니다. 확인해주세요') 
                        #     else:
                        #         driver.find_element(by=By.XPATH, value=xpath1).click()
                        # else:
                        #     driver.find_element(by=By.XPATH, value='/html/body/div[4]/div[2]/div[2]/div[2]/div/div[1]/button').click()
                        #     self.returnMessage.emit(f'{str(df_input[df_input.columns[7]][i]).strip()} 가 부품마스터에 등록되지 않았습니다. 확인해주세요')        
                        
                        time.sleep(1.5) 
                        driver.find_element(by=By.XPATH, value='//input[@name="eserv__Last_Inventory_Dt__c"]').send_keys(df_input[df_input.columns[27]][i].strftime('%Y. %m. %d.'))
                        time.sleep(0.5)
                        driver.find_element(by=By.XPATH, value='//button[@name="SaveEdit"]').click()
                        time.sleep(2)
                        xpath_help = f'//div[text()=""부품", "사이트", "창고" 혹은 "선반"이 중복되었습니다."]'
                        if self.hasXpath(driver, xpath_help):
                            self.returnMessage.emit(f'「{str(df_input[df_input.columns[7]][i])}」가 부품, 사이트, 창고 혹은 선반이 중복되어 재고등록이 불가합니다. 재고미등록 리스트는 Output폴더를 확인해주세요.')
                            df_output = df_output.append(df_input.iloc[i])   
                            df_output = df_output.reset_index(drop=True)
                            df_output.to_excel('.\\Output\\재고미등록리스트_'+str(today)+'.xlsx')
                        break
                self.returnI.emit(df_input.index[-1] + i)
            self.returnEnd.emit(True)
            self.thread().quit()
        except Exception as e:
            self.returnError.emit(e)
            self.thread().quit()
            return

class QTextEditLogger(logging.Handler):
    def __init__(self, parent):
        super().__init__()
        self.widget = QPlainTextEdit(parent)
        self.widget.setGeometry(QRect(10, 260, 661, 161))
        self.widget.setReadOnly(True)
        self.widget.setPlainText('')
        self.widget.setStyleSheet('background-color: rgb(53, 53, 53);\ncolor: rgb(255, 255, 255);')
        self.widget.setObjectName('logBrowser')
        font = QFont()
        font.setFamily('Nanum Gothic')
        font.setBold(False)
        font.setPointSize(9)
        self.widget.setFont(font)

    def emit(self, record):
        msg = self.format(record)
        self.widget.appendPlainText(msg)

class Ui_MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi()

    def setupUi(self):
        logger = logging.getLogger(__name__)
        rfh = RotatingFileHandler(filename='./Log.log', 
                                    mode='a',
                                    maxBytes=5*1024*1024,
                                    backupCount=2,
                                    encoding=None,
                                    delay=0
                                    )
        logging.basicConfig(level=logging.DEBUG, 
                            format = '%(asctime)s:%(levelname)s:%(message)s', 
                            datefmt = '%m/%d/%Y %H:%M:%S',
                            handlers=[rfh])
        self.setObjectName('MainWindow')
        self.resize(700, 600)
        self.setStyleSheet('background-color: rgb(252, 252, 252);')
        self.centralwidget = QWidget(self)
        self.centralwidget.setObjectName('centralwidget')
        self.gridLayout2 = QGridLayout(self.centralwidget)
        self.gridLayout2.setObjectName('gridLayout2')
        self.gridLayout = QGridLayout()
        self.gridLayout.setObjectName('gridLayout')
        self.groupBox = QGroupBox(self.centralwidget)
        self.groupBox.setTitle('')
        self.groupBox.setObjectName('groupBox')
        self.gridLayout4 = QGridLayout(self.groupBox)
        self.gridLayout4.setObjectName('gridLayout4')
        self.gridLayout3 = QGridLayout()
        self.gridLayout3.setObjectName('gridLayout3')
        self.le_userName = QLineEdit(self.groupBox)
        self.le_userName.setMinimumSize(QSize(0, 25))
        self.le_userName.setObjectName('le_userName')
        self.gridLayout3.addWidget(self.le_userName, 0, 1, 1, 1)
        self.le_password = QLineEdit(self.groupBox)
        self.le_password.setMinimumSize(QSize(0, 25))
        self.le_password.setObjectName('le_password')
        self.le_password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.gridLayout3.addWidget(self.le_password, 1, 1, 1, 1)
        self.btn_fileSelect = QToolButton(self.groupBox)
        self.btn_fileSelect.setMinimumSize(QSize(0,25))
        self.btn_fileSelect.setObjectName('btn_fileSelect')
        self.gridLayout3.addWidget(self.btn_fileSelect, 3, 1, 1, 1)

        self.label_filePath = QLabel(self.groupBox)
        self.label_filePath.setAlignment(Qt.AlignLeft | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label_filePath.setObjectName('label_filePath')
        self.gridLayout3.addWidget(self.label_filePath, 4, 1, 1, 1)
       
        self.label_blank = QLabel(self.groupBox)
        self.label_blank.setObjectName('label_blank')
        self.gridLayout3.addWidget(self.label_blank, 7, 4, 1, 1)

        self.line2 = QFrame(self.groupBox)
        self.line2.setFrameShape(QFrame.HLine)
        self.line2.setFrameShadow(QFrame.Sunken)
        self.line2.setObjectName('line2')
        self.gridLayout3.addWidget(self.line2, 6, 0, 1, 10)
        self.progressBar = QProgressBar(self.groupBox)
        self.progressBar.setObjectName('progressBar')
        self.gridLayout3.addWidget(self.progressBar, 7, 1, 1, 2)
        self.btn_run = QToolButton(self.groupBox)
        sizePolicy = QSizePolicy(QSizePolicy.Ignored, 
                                    QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.btn_run.sizePolicy().hasHeightForWidth())
        self.btn_run.setSizePolicy(sizePolicy)
        self.btn_run.setMinimumSize(QSize(80, 35))
        self.btn_run.setStyleSheet('background-color: rgb(63, 63, 63);\ncolor: rgb(255, 255, 255);')
        self.btn_run.setObjectName('btn_run')
        self.gridLayout3.addWidget(self.btn_run, 7, 3, 1, 1)
        self.label = QLabel(self.groupBox)
        self.label.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label.setObjectName('label')
        self.gridLayout3.addWidget(self.label, 0, 0, 1, 1)
        self.label_password = QLabel(self.groupBox)
        self.label_password.setAlignment(Qt.AlignRight | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label_password.setObjectName('label_password')
        self.gridLayout3.addWidget(self.label_password, 1, 0, 1, 1)

        self.label_fileSelect = QLabel(self.groupBox)
        self.label_fileSelect.setAlignment(Qt.AlignLeft | 
                                Qt.AlignTrailing | 
                                Qt.AlignVCenter)
        self.label_fileSelect.setObjectName('label_fileSelect')
        self.gridLayout3.addWidget(self.label_fileSelect, 3, 0, 1, 1) 
        self.line = QFrame(self.groupBox)
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)
        self.line.setObjectName('line')
        self.gridLayout3.addWidget(self.line, 8, 0, 1, 10)
        self.gridLayout4.addLayout(self.gridLayout3, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox, 0, 0, 1, 1)
        self.groupBox2 = QGroupBox(self.centralwidget)
        self.groupBox2.setTitle('')
        self.groupBox2.setObjectName('groupBox2')
        self.gridLayout6 = QGridLayout(self.groupBox2)
        self.gridLayout6.setObjectName('gridLayout6')
        self.gridLayout5 = QGridLayout()
        self.gridLayout5.setObjectName('gridLayout5')
        self.logBrowser = QTextEditLogger(self.groupBox2)
        self.logBrowser.setFormatter(
                                    logging.Formatter('[%(asctime)s] %(levelname)s:%(message)s', 
                                                        datefmt='%Y-%m-%d %H:%M:%S')
                                    )
        logging.getLogger().addHandler(self.logBrowser)
        logging.getLogger().setLevel(logging.INFO)
        self.gridLayout5.addWidget(self.logBrowser.widget, 0, 0, 1, 1)
        self.gridLayout6.addLayout(self.gridLayout5, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.groupBox2, 1, 0, 1, 1)
        self.gridLayout2.addLayout(self.gridLayout, 0, 0, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(self)
        self.menubar.setGeometry(QRect(0, 0, 653, 21))
        self.menubar.setObjectName('menubar')
        self.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(self)
        self.statusbar.setObjectName('statusbar')
        self.setStatusBar(self.statusbar)
        self.retranslateUi(self)

        self.btn_fileSelect.clicked.connect(self.importFile)
        self.btn_run.clicked.connect(self.run)
        #디버그용 플래그
        self.isDebug = True
        if self.isDebug:
            self.le_userName.setText('sungik.jang@ymfk.yokogawa.com')
            self.le_password.setText('ymfk2345')

        self.thread = QThread()
        self.thread.setTerminationEnabled(True)
        self.show()

    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        MainWindow.setWindowTitle(_translate('MainWindow', 'eServ 자동 등록 프로그램 Rev0.00'))
        MainWindow.setWindowIcon(QIcon('.\\Logo\\logo.png'))
        self.label.setText(_translate('MainWindow', 'eServ UserName:'))
        self.label_password.setText(_translate('MainWindow', 'eServ PassWord:'))
        self.btn_run.setText(_translate('MainWindow', '실행'))
        self.label_fileSelect.setText(_translate('MainWndow', '부품 등록 파일 입력 :'))
        self.btn_fileSelect.setText(_translate('MainWindow', ' 파일 선택 '))
        self.label.setText(_translate('MainWindow', 'eServ UserName:'))
        self.label_blank.setText(_translate('MainWindow', ' 사이즈'))
        self.label_filePath.setText(_translate('MainWindow', '파일 미지정'))
        logging.info('프로그램이 정상 기동했습니다')
    
    @pyqtSlot()
    def importFile(self):
        try:
            fileName = QFileDialog.getOpenFileName(self, 'Open File', './', 'Excel Files (*.xlsx)')[0]
            if fileName != "":
                self.label_filePath.setText(fileName)
                logging.info('파일 불러오기 완료')
        except Exception as e:
            logging.info('파일 불러오기 실패')

    def setTotalCnt(self, totalCnt):
        self.progressBar.setRange(0, totalCnt)

    def showError(self, e):
        logging.error(e)
        self.btn_run.setEnabled(True)
    def searchEnd(self, Flag):
        if Flag:
            logging.info('부품 마스터 등록 및 재고 등록이 완료되었습니다.')
        self.btn_run.setEnabled(True)

    def showMessage(self, str):
        logging.warning(str)

    def run(self):
        if len(self.le_userName.text()) > 0:
            if len(self.le_password.text()) > 0:
                if self.label_filePath.text() != '파일 미지정':
                    self.btn_run.setEnabled(False)
                    self.x = ThreadClass(self.le_userName.text(), self.le_password.text(), self.label_filePath.text())
                    self.x.moveToThread(self.thread)
                    self.thread.started.connect(self.x.run)
                    self.x.returnTotalCnt.connect(self.setTotalCnt)
                    self.x.returnI.connect(self.progressBar.setValue)
                    self.x.returnError.connect(self.showError)
                    self.x.returnEnd.connect(self.searchEnd)
                    self.x.returnMessage.connect(self.showMessage)
                    self.thread.start()
                else:
                    logging.warning('지정된 파일이 없습니다. 다시 한 번 확인해주세요')
            else:
                logging.warning('Password가 입력되지 않았습니다. 다시 한 번 확인해주세요')
        else:
            logging.warning('Username이 입력되지 않았습니다. 다시 한 번 확인해주세요')
if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ui = Ui_MainWindow()
    sys.exit(app.exec_())


