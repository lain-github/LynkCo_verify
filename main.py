import time
# import webbrowser
# import cv2
import random
import openpyxl
# from urllib import request
from selenium import webdriver
# import numpy as np
from selenium.webdriver.common.action_chains import ActionChains
import base64
from PIL import Image

# chrome_options = webdriver.ChromeOptions()
# chrome_options.add_argument(r'--user-data-dir=C:\Users\GL\AppData\Local\Google\Chrome\User Data\Default')
# chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
# chrome_driver = r'C:\\Program Files\\Google\\Chrome\\Application\\chromedriver.exe'
# browser = webdriver.Chrome(executable_path=chrome_driver, options=chrome_options)

chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("excludeSwitches", ['enable-automation'])
chrome_driver = r'C:\\Program Files\\Google\\Chrome\\Application\\chromedriver.exe'
browser = webdriver.Chrome(executable_path=chrome_driver, options=chrome_options)

def move(btn, y):
    distance = y
    has_gone_dist = 0
    remaining_dist = y
    # 获取滑块
    element = browser.find_element_by_xpath(btn)
    ActionChains(browser).click_and_hold(on_element=element).perform()
    time.sleep(0.5)

    while remaining_dist > 0:
        ratio = remaining_dist / distance
        if ratio < 0.3:
            span = random.randint(5, 8)
        elif ratio > 0.9:
            span = random.randint(9, 15)
        else:
            span = random.randint(20, 25)

        ActionChains(browser).move_by_offset(span, random.randint(-5, 5)).perform()
        remaining_dist -= span
        has_gone_dist += span
        time.sleep(random.randint(5, 20) / 100)
    ActionChains(browser).move_by_offset(remaining_dist, random.randint(-5, 5)).perform()
    ActionChains(browser).release(on_element=element).perform()
    time.sleep(2)


def is_similar_color(x_pixel, y_pixel):
    for i, pixel in enumerate(x_pixel):
        if abs(y_pixel[i] - pixel) > 90:
            return False
    return True


def get_offset_distance(cut_image, full_image):
    print(cut_image.width, cut_image.height)
    for x in range(cut_image.width):
        for y in range(cut_image.height):
            cpx = cut_image.getpixel((x, y))
            fpx = full_image.getpixel((x, y))
            if not is_similar_color(cpx, fpx):
                img = cut_image.crop((x, y, x + 50, y + 40))
                # 保存一下计算出来位置图片，看看是不是缺口部分
                img.save("1.png")
                return x

def getpic():
    time.sleep(1)
    # 保存拼图
    backimg = "backimg.jpg"
    slideimg = "slideimg.jpg"

    # 下面的js代码根据canvas文档说明而来
    JS = 'return document.getElementsByClassName("geetest_canvas_bg geetest_absolute")[0].toDataURL("image/png");'
    # 执行 JS 代码并拿到图片 base64 数据
    im_info = browser.execute_script(JS)  # 执行js文件得到带图片信息的图片数据
    im_base64 = im_info.split(',')[1]  # 拿到base64编码的图片信息
    im_bytes = base64.b64decode(im_base64)  # 转为bytes类型
    with open(backimg, 'wb') as f:  # 保存图片到本地
        f.write(im_bytes)

    # 下面的js代码根据canvas文档说明而来
    JS = 'return document.getElementsByClassName("geetest_canvas_fullbg")[0].toDataURL("image/png");'
    # 执行 JS 代码并拿到图片 base64 数据
    im_info = browser.execute_script(JS)  # 执行js文件得到带图片信息的图片数据
    im_base64 = im_info.split(',')[1]  # 拿到base64编码的图片信息
    im_bytes = base64.b64decode(im_base64)  # 转为bytes类型
    with open("full.jpg", 'wb') as f:  # 保存图片到本地
        f.write(im_bytes)

    JS = 'return document.getElementsByClassName("geetest_canvas_slice geetest_absolute")[0].toDataURL("image/png");'
    # 执行 JS 代码并拿到图片 base64 数据
    im_info = browser.execute_script(JS)  # 执行js文件得到带图片信息的图片数据
    im_base64 = im_info.split(',')[1]  # 拿到base64编码的图片信息
    im_bytes = base64.b64decode(im_base64)  # 转为bytes类型
    with open(slideimg, 'wb') as f:  # 保存图片到本地
        f.write(im_bytes)
    # 完毕

    cut_image = Image.open(backimg)
    full_image = Image.open("full.jpg")

    x = get_offset_distance(cut_image, full_image)
    print(x)
    # btn = '//div[@class="geetest_slider_button"]'
    # move(btn, x - 6)
    #
    #
    #
    # # 读入灰度图
    # block = cv2.imread(slideimg, 0)
    # template = cv2.imread(backimg, 0)
    # # 保存灰度图
    # blockname = "block.jpg"
    # templatename = "template.jpg"
    # cv2.imwrite(blockname, block)
    # cv2.imwrite(templatename, template)
    #
    # contours, hierarchy = cv2.findContours(block, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    # image1 = block.copy()
    # for item in contours:
    #     rect = cv2.boundingRect(item)
    #     x = rect[0]
    #     y = rect[1]
    #     weight = rect[2]
    #     height = rect[3]
    #     cv2.imwrite(blockname, image1[y:y + height, x:x + weight])
    #     break
    # cv2.imshow("before", block)
    # # 滑块再处理
    # block = cv2.imread(blockname)
    # block = cv2.cvtColor(block, cv2.COLOR_RGB2GRAY)
    # # block1 = block.copy()
    # retval, dst = cv2.threshold(block, 10, 200, cv2.THRESH_BINARY)
    # # block = cv2.GaussianBlur(block, (3, 3), 1)
    # # overlapping = cv2.addWeighted(dst, 1, block, 0.5, 0)
    #
    # cv2.imwrite(blockname, dst)
    #
    # # overlapping = abs(255 - block)
    # # cv2.imwrite(blockname, overlapping)
    #
    # block = cv2.imread(blockname)
    # template = cv2.imread(templatename)
    # retval, dst = cv2.threshold(template, 140, 200, cv2.THRESH_BINARY)
    # result = cv2.matchTemplate(block, dst, cv2.TM_CCOEFF_NORMED)
    # min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
    # print('min_val:', min_val, 'max_val:', max_val)
    # print('**********', min_loc)
    # x, y = min_loc
    # # x, y = np.unravel_index(result.argmax(), result.shape)
    # print("x is {} y is {}".format(x, y))
    # print("x方向的偏移", int(y / 1.3), 'x:', x, 'y:', y)
    # cv2.imshow("template", template)
    # cv2.imshow("block", block)
    # cv2.waitKey()
    # btn = '//div[@class="geetest_slider_button"]'
    # move(btn, x, int(y))
    # move(btn, x, y)


def getpersoninfo(channel):
    """
    :param channel:
    :return:
    """
    wb = openpyxl.load_workbook('personinfo.xlsx')
    sheet = wb[channel]
    url = sheet['A2'].value
    username = sheet['B2'].value
    pwd = sheet['C2'].value
    return url, username, pwd


def log_in():
    url = 'https://www.lynkco.com.cn/'
    username = 'account'
    password = 'password'

    browser.get(url)
    browser.implicitly_wait(4)
    agree = browser.find_element_by_xpath("//div/span[contains(text(),'同意')]")
    browser.execute_script("arguments[0].click();", agree)
    browser.maximize_window()
    link_login = browser.find_element_by_id('btnLoginPC')
    link_login.click()
    time.sleep(5)
    access = r'//div/div[@class="svg-btn"]'
    link_login = browser.find_element_by_xpath(access)
    link_login.click()
    link_login = browser.find_element_by_link_text('密码登录')
    link_login.click()
    time.sleep(1)
    user = browser.find_element_by_id('login-by')
    user.clear()
    pwd = browser.find_element_by_id('password')
    pwd.clear()
    time.sleep(1)
    user.send_keys(username)
    time.sleep(1)
    pwd.send_keys(password)
    time.sleep(1)

    verify = r'//div/div[@class="geetest_radar_tip"]'
    link_login = browser.find_element_by_xpath(verify)
    link_login.click()


def main(name):
    print(f'Hi, {name}')
    log_in()

    # TODO: getpic and slide identify
    while True:
        try:
            getpic()
        except:
            print('登录成功')
            break


if __name__ == '__main__':
    main('PyCharm')
