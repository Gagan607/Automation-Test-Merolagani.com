import pandas as pd

from selenium import webdriver
import time
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import excel_operation

import send_email

test_case_location = "test_case/test_case.xlsx"
test_result_location = "result/test_result/test_result.xlsx"


def read_excel():
    reader = pd.read_excel(test_case_location)
    for row, column in reader.iterrows():
        sn = column["SN"]
        test_summary = column["Test_Summary"]
        xpath = column["Xpath"]
        action = column["Action"]
        value = column["Value"]
        action_defination(sn, test_summary, xpath, action, value)


def action_defination(sn, test_summary, xpath, action, value):

    if action == 'open_browser':
        result, remarks = open_browser(value)
    elif action == 'open_url':
        result, remarks = open_url(value)
    elif action == 'close_browser':
        result, remarks = close_browser()
    elif action == 'click':
        result, remarks = click(xpath)
    elif action == 'verify_text':
        result, remarks = verify_text(xpath, value)
    elif action == 'input_text':
        result, remarks = input_text(xpath, value)
    elif action == 'select_dropdown':
        result, remarks = select_dropdown(xpath, value)
    elif action == 'wait':
        result, remarks = wait(value)
    elif action == 'verify_alert_text':
        result, remarks = verify_alert_text(xpath, value)
    else:
        result, remarks = "FAIL", "Action not defined"
    print(sn, test_summary, result, remarks)
    excel_operation.write_result(sn, test_summary, result, remarks)
    excel_operation.write_result2(sn, result, remarks)


def open_browser(value):
    try:
        global driver
        if value == 'chrome':
            s = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=s)
            driver.maximize_window()
            result = "PASS"
            remarks = ""
        elif value == 'firefox':
            # firefox code here
            result = "PASS"
            remarks = ""
        else:
            result = "FAIL"
            remarks = "value ,Browser not supported"
    except Exception as ex:
        result = "FAIL"
        remarks = ex
    return result, remarks


def open_url(value):
    try:
        driver.get(value)
        result = "PASS"
        remarks = ""
    except Exception as ex:
        result = "FAIL"
        remarks = ex
    return result, remarks


def close_browser():
    try:
        driver.quit()
        result = "PASS"
        remarks = ""
    except Exception as ex:
        result = "FAIL"
        remarks = ex
    return result, remarks


def click(xpath):
    try:
        driver.find_element(By.XPATH, xpath).click()
        result = "PASS"
        remarks = ""

    except Exception as ex:
        result = "FAIL"
        remarks = ex
    return result, remarks


def verify_text(xpath, value):
    try:
        actual_text = driver.find_element(By.XPATH, xpath).text
        try:
            assert actual_text == value
        except AssertionError:
            result = "FAIL"
            remarks = "Text don't match, actual text: " + actual_text + "expected test: " + value
        else:
            result = "PASS"
            remarks = ""
    except Exception as ex:
        result = "FAIL"
        remarks = ex
    return result, remarks


def verify_alert_text(xpath, value):
    try:
        driver.find_element(By.XPATH, xpath).send_keys(value)
        result = "PASS"
        remarks = ""
    except Exception as ex:
        result = "FAIL"
        remarks = "Can be verified only when empty field."
    return result, remarks


def input_text(xpath, value):
    try:
        driver.find_element(By.XPATH, xpath).send_keys(value)
        result = "PASS"
        remarks = ""
    except Exception as ex:
        result = "FAIL"
        remarks = ex
    return result, remarks


def select_dropdown(xpath, value):
    try:
        dropdown_xpath = driver.find_element(By.XPATH, xpath)
        menu = Select(dropdown_xpath)
        menu.select_by_visible_text(value)
        result = "PASS"
        remarks = ""
    except Exception as ex:
        result = "FAIL"
        remarks = ex
    return result, remarks


def wait(value):
    try:
        time.sleep(value)
        result = "PASS"
        remarks = ""

    except Exception as ex:
        result = "FAIL"
        remarks = ex
    return result, remarks


if __name__ == "__main__":
    excel_operation.clear_result()
    excel_operation.write_header()
    excel_operation.write_header1()
    excel_operation.write_summary()
    read_excel()
    # send_email.send_selenium_report()