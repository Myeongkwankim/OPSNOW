from selenium.common.exceptions import NoSuchElementException
import glob, os
import time
from datetime import date
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pyodbc
import pandas as pd
from decimal import Decimal
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.styles import Font
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import date
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def check_exist_by_xpath(element, xpath):
    try:
        element.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return  False
    return True

def select_languge(languge):
    print("언어 선택 : {0}".format(languge))

    driver.find_element_by_xpath('.//div[@id="language"]/button').click()
    # English
    languagelist = driver.find_elements_by_xpath('.//div[@id="language"]/div/ul/li')
    for item in languagelist:
        if item.text == languge:
            item.click()
            time.sleep(5)
            break

def Select_company(company_name):
    driver.find_element_by_xpath('//button[@class="btn-companies"]').click()
    time.sleep(1)
    if check_exist_by_xpath(driver,'//input[@id="company_search_word"]'):
        driver.find_element_by_xpath('//input[@id="company_search_word"]').send_keys(company_name)
    else:
        driver.find_element_by_xpath('//fieldset[@class="search-word"]/input').send_keys(company_name)

    companys = driver.find_elements_by_xpath('//ul[@class="list-companies"]/li')

    for company in companys:
        if company.text == company_name:
            company.click()

def Get_Company_list():
    driver.find_element_by_xpath('//button[@class="btn-companies"]').click()
    time.sleep(2)

    list = driver.find_elements_by_xpath('.//ul[@class="list-companies"]/li')
    for company in list:
        company_name = company.text
        get_company_qry = "exec prc_get_company_info '{0}'".format(company_name)
        data = pd.read_sql(get_company_qry, cnxn).values

        if len(data) == 0:
            set_company_qry = "exec prc_set_company_info '{0}'".format(company_name)
            cursor.execute(set_company_qry)
            cnxn.commit()

cnxn = pyodbc.connect("Driver={SQL Server Native Client 11.0};"
                      "Server=10.30.220.96;"
                      "Database=Bsp_Management_WhiteLabel;"
                      "uid=qateam;pwd=bespin!2018")
cursor = cnxn.cursor()

# Main : Login
driver = webdriver.Chrome('D:/chromedriver.exe')
driver .set_window_size(1920,1400)
driver.get("https://metering.sec-ecm.net/")
#OPSNOW Login
time.sleep(2)
driver.find_element_by_id('username').send_keys('jihwan.park@bespinglobal.com')
time.sleep(1)
driver.find_element_by_id('password').send_keys('qkrwlghks1!')
time.sleep(2)
driver.find_element_by_xpath('//*[@class="btn-login"]').click()
time.sleep(5)

Get_Company_list()