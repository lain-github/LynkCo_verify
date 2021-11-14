import time
import cv2
import random
import openpyxl
from urllib import request
from selenium import webdriver
import numpy as np
from selenium.webdriver.common.action_chains import ActionChains

chrome_driver = r'C:\\Program Files\\Google\\Chrome\\Application\chromedriver.exe'
browser = webdriver.chrome(executable_path=chrome_driver)

def getpersoninfo(channel):
    wb = openpyxl.load_workbook('personinfo.xlsx')
    sheet = wb[channel]
    url = sheet['A2'].value
    username = sheet['B2'].value
    pwd = sheet['C2'].value
    return url, username, pwd

def loginjd():
    channel = 'jd'
    url = getpersoninfo(channel)[0]
    username = getpersoninfo(channel)[1]
    password = getpersoninfo(channel)[2]
    browser.get(url)
    browser.maximize_window()
    link_login = browser.find_element_by_link_text('你好，请登录')
    link_login.click()

def main(name):
    print(f'Hi, {name}')
    loginjd()

if __name__ == '__main__':
    main('PyCharm')

