from selenium.common.exceptions import NoSuchElementException
import glob, os
import time
from datetime import date
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pyodbc
import pandas as pd
import sys
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

def Select_Service(service_name):
    #현재 선택되어 있는 서비스 - 우측 선택 리스트 클릭
    driver.find_element_by_xpath('.//button[@class="btn-selected-service"]').click()
    #service name 선택하여
    time.sleep(2)

    if check_exist_by_xpath(driver,'.//div[@class="service-container-inner"]/ul[@class="list-service"]/li[@name="{0}"]'.format(service_name)):
        service_list = driver.find_element_by_xpath('.//div[@class="service-container-inner"]/ul[@class="list-service"]/li[@name="{0}"]'.format(service_name))
        service_list.click()
    else:
    #print(len(service_list))
        if service_name == "menu_asset":
            service_name = "Asset Management"

        service_list = driver.find_elements_by_xpath('.//ul[@class="list-service"]/li')
        for service in service_list:
            print(service.text)
            if service.text == service_name:
                print("ok")
                service.click()
                break;

def Click_Menu(menu_name):
    result = 0
    menulist = driver.find_elements_by_xpath('//div[@class="submenus-container"]/ul/li')
    for menu in menulist:
        if menu.text == menu_name:
            menu.click()
            result = 1

    return  result

def cost_dashboard_check(company_name):
    time.sleep(10)
    dashboard_vendors = driver.find_elements_by_xpath(
        './/div[@class="dashboard-item vendor"]/div[@class="dashboard-item-box"]/label')
    for dv in dashboard_vendors:
        dv.click()
        time.sleep(10)

        time.sleep(15)
        driver.execute_script("window.scrollTo(0,1300);")
        body_contain = driver.find_element_by_xpath(
            './/section[@class="dashboard-section item-cost"]/div[@id="item-cost"]')
        # tab list
        tabs = body_contain.find_elements_by_xpath('.//div[@class="common-tabs"]/button')
        for tab in tabs:
            print("Dashboard : itemized Billing Amount - {0}".format(tab.text))
            tab.click()
            time.sleep(20)
            check_str = "exec Metering_samsung.dbo.prc_Cost_AutoTesting_Daily_log '{0}','{1}','{2}','{3}','{4}'"\
                .format(company_name,"Dashboard","{0}:{1}".format(dv.text,tab.text),"Check","")
            print(check_str)
            cursor.execute(check_str)
            cnxn.commit()


def cost_billingAnalytics_check(company_name):
    vendor_cnt = 0
    vendors = driver.find_elements_by_xpath('.//div[@class="qs-items vendor"]/p')
    while len(vendors) > vendor_cnt:
        vendor = driver.find_elements_by_xpath('.//div[@class="qs-items vendor"]/p')[vendor_cnt]
        vendor_name = vendor.text

        if vendor_cnt > 0:
            # Billing Analytics 화면에서 Intelligent search 화면으로 전환
            time.sleep(5)
            driver.find_element_by_xpath('.//button[@class="button-normal icon search"]').click()
            time.sleep(20)

        vendor.click()
        time.sleep(10)
        # Apply button click
        driver.find_elements_by_xpath('.//button[@class="button-normal icon search"]')[0].click()
        time.sleep(20)
        type_tabs = driver.find_elements_by_xpath('.//div[@class="common-tabs"]/button')
        for tab in type_tabs:
            print("billing Analytics - {0} Tab click".format(tab.text))
            tab.click()
            time.sleep(30)
            check_str = "exec Metering_samsung.dbo.prc_Cost_AutoTesting_Daily_log '{0}','{1}','{2}','{3}','{4}'" \
                .format(company_name, "Dashboard", "{0}:{1}".format(vendor_name, tab.text), "Check", "")
            cursor.execute(check_str)
            cnxn.commit()

        vendor_cnt += 1


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
driver.find_element_by_id('username').send_keys('hyeokjin.seong@bespinglobal.com')
time.sleep(1)
driver.find_element_by_id('password').send_keys('tjdgurwls!3')
time.sleep(2)
driver.find_element_by_xpath('//*[@class="btn-login"]').click()
time.sleep(5)

company_str = "Select company_name from [Bsp_Management_WhiteLabel].[dbo].[TB_Company_Info] Where status_Metering = 1 and isDeleted = 0"
company_list = pd.read_sql(company_str, cnxn).values
print(company_list)
for company_ in company_list:
    company = company_[0]

    try:
        # Main : check function
        ## 1. Dashboard - item Grid tabs checking
        Select_company(company)
        time.sleep(20)
        print(company)
        Click_Menu("Dashboard")
        cost_dashboard_check(company)

        ischeck =  Click_Menu("Billing Analytics")
        time.sleep(20)
        if ischeck == 1:
            cost_billingAnalytics_check(company)
    except:
        print("Unexpected error:", sys.exc_info()[0])

driver.close()