import sys
import os
from ctypes import *

from selenium import webdriver
import time
import csv

from PIL import Image

import multiprocessing
import threading

def yanzhengma(n):

    YDMApi = windll.LoadLibrary('yundamaAPI-x64.dll')
    appId = 1
    appKey = b'22cc5376925e9387a23cf797cb9ba745'
    username = b''#''里面是你的云打码账号
    password = b''#''里面是你的密码
    codetype = 1008#这里是验证码的类别，1008是8位的验证码
    result = c_char_p(b"                              ")
    timeout = 60

    balance = YDMApi.YDM_GetBalance(username, password)
    print('用户名：%s，剩余题分：%d' % (username, balance))
    filename = 'D:\luvm\onedrive\python_class\work\grade_mul\pic\pic{}.png'.format(n).encode('utf-8')
    #这里是文件存储的路径，你可以把解压文件夹"grade_mul"中“pic”文件夹所在的路径覆盖“pic{}.png”前面的部分
    print('云端打码中')
    captchaId = YDMApi.YDM_EasyDecodeByPath(username, password, appId, appKey, filename, codetype, timeout, result)
    print("一键识别：验证码ID：%d，识别结果：%s" % (captchaId, result.value))
    print(str(result.value)[2:-1])
    if len(str(result.value)[2:-1]) != 8:
        re = YDMApi.YDM_EasyReport(username, password, appId, appKey,captchaId,False)
        print('识别错误,处理结果：',re)
        return yanzhengma(n)
    return str(result.value)[2:-1]

def csv_saver(index,content):
    print('获取到数据')
    with open('D:\luvm\onedrive\python_class\work\grade_mul\data.csv', 'at', newline='', encoding='gbk') as f:
        # 这里是文件存储的路径，你可以把解压文件夹"grade_mul"所在的路径覆盖“data.csv”前面的部分
        writer = csv.writer(f)
        writer.writerow(index)
        writer.writerow(content)

def get_data(n):
    try:
        index = []
        content = []
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('-ignore-certificate-errors')
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.binary_location = (r'C:\Users\jack8\AppData\Local\Google\Chrome\Application\chrome.exe')#这里是chrome浏览器的路径，一般把“jack8”替换成你的用户名就可以了
        driver = webdriver.Chrome(options=chrome_options)

        driver.get('http://zsb.jlu.edu.cn/Index/graduate.html')
        driver.set_window_size(1280, 800)
        driver.find_element_by_id("IDNumber").clear()
        driver.find_element_by_id("IDNumber").send_keys("0")
        driver.find_element_by_id("Admission").clear()
        if n <= 9:
            mid_num = '0000'
        elif n <= 99:
            mid_num = '000'
        elif n <= 999:
            mid_num = '00'
        elif n <= 9999:
            mid_num = '0'
        else:
            mid_num = ''
        num = '1018392142' + mid_num + str(n)
        driver.find_element_by_id("Admission").send_keys(num)

        driver.get_screenshot_as_file('D:\luvm\onedrive\python_class\work\grade_mul\pic\screen{}.png'.format(n))
        # 这里是文件存储的路径，你可以把解压文件夹"grade_mul"中“pic”文件夹所在的路径覆盖“screen{}.png”前面的部分
        im = Image.open("D:\luvm\onedrive\python_class\work\grade_mul\pic\screen{}.png".format(n))
        box = (650, 480, 905, 520)
        cropImg = im.crop(box)
        cropImg.save("D:\luvm\onedrive\python_class\work\grade_mul\pic\pic{}.png".format(n))
        # 这里是文件读取的路径，你可以把解压文件夹"grade_mul"中“pic”文件夹所在的路径覆盖“pic{}.png”前面的部分
        yzm = yanzhengma(n)
        driver.find_element_by_id('yzm').send_keys(yzm)
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/section[3]/div/div/form/div[5]/div/button').click()
        time.sleep(1)
        try:
            index.append(driver.find_element_by_xpath('/html/body/section[3]/div/div/table/tbody/tr[1]/td/span').text)
            index.append(driver.find_element_by_xpath('/html/body/section[3]/div/div/table/tbody/tr[2]/td[1]/span').text)
            index.append(driver.find_element_by_xpath('/html/body/section[3]/div/div/table/tbody/tr[2]/td[2]/span').text)
            index.append(driver.find_element_by_xpath('/html/body/section[3]/div/div/table/tbody/tr[2]/td[3]/span').text)
            index.append(driver.find_element_by_xpath('/html/body/section[3]/div/div/table/tbody/tr[2]/td[4]/span').text)
            index.append(driver.find_element_by_xpath('/html/body/section[3]/div/div/table/tbody/tr[2]/td[5]/span').text)
            content.append('#')
            content.append(driver.find_element_by_xpath('/html/body/section[3]/div/div/table/tbody/tr[3]/td[1]/span').text)
            content.append(driver.find_element_by_xpath('/html/body/section[3]/div/div/table/tbody/tr[3]/td[2]/span').text)
            content.append(driver.find_element_by_xpath('/html/body/section[3]/div/div/table/tbody/tr[3]/td[3]/span').text)
            content.append(driver.find_element_by_xpath('/html/body/section[3]/div/div/table/tbody/tr[3]/td[4]/span').text)
            content.append(driver.find_element_by_xpath('/html/body/section[3]/div/div/table/tbody/tr[3]/td[5]/span').text)
            csv_saver(index, content)
        except:
            pass
        driver.close()
    except:
        get_data(n)


if __name__ == '__main__':
    n = int(input('输入起始值:'))
    while True:
        ps = multiprocessing.Pool(10)
        li =[]
        for i in range(100):
            li.append(n)
            n -= 1
        print(li)
        ps.map(get_data, li)
        ps.close()
        ps.join()
        if n < 0:
            break


