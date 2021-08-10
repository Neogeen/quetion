from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
import timeit as t
import pandas as pd
from selenium import webdriver
from docxtpl import DocxTemplate
import io
from PIL import Image
from docx.shared import Inches

# Простоянные переменные
driver = webdriver.Firefox()
app_login = 'univer@kaznu.kz'
app_password = 'yhxTQRBc2n'
delay = 9


driver.get('https://app.oqylyq.kz/')
sleep(1.7)
driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div/div/div/div[2]/div[2]/div/form/div[1]/div[1]/div/input').send_keys(app_login)
driver.find_element_by_xpath('//*[@id="__layout"]/div/div[1]/div/div/div/div[2]/div[2]/div/form/div[2]/div/div/input').send_keys(app_password)
driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div/div/div/div[2]/div[2]/div/form/div[4]/div/button/div[2]').click()
sleep(3)
i = 2
# for i in range(1, 23):
driver.get(f'https://app.oqylyq.kz/t/assignments?page=2')
sleep(1)

for j in range(2,10):
    dl = 5
    if j == 11:
        nope = 0
        break
    sleep(2)
    exz = str(driver.find_element_by_xpath(f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[2]/div/a[{j}]/div[2]/div[1]/div[1]').text)
    print(f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[2]/div/a[{j}]/div[1]/div/div/div/div[2]/div/div')
    c = driver.find_element_by_xpath(f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[2]/div/a[{j}]/div[1]/div/div/div/div[2]/div/div').text
    if c == '0':
        # driver.get(f'https://app.oqylyq.kz/t/assignments?page={i}')
        continue
    sleep(1)


    driver.find_element_by_xpath(f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[2]/div/a[{j}]/div[2]/div[1]/div[1]').click()


    sleep(2)
    # driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/a').click()
    student_number = 1
    for i in range(130):
        bl = 1
        rp = 0
        pl = 0
        while True:
            try:
                driver.find_element_by_xpath(f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/a[{student_number}]/div[2]/div').click()
                break
            except:
                rp += 1
                if rp > 11:
                    break
                continue
        if rp > 11:
            break
        while True:
            try:
                driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[3]/div[2]/span[1]/div').click()
                #/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[3]/div[2]/span[1]/div
                break
            except:
                continue

        question3 = ""
        question4 = ""
        sleep(1)
        question1 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[2]/div').screenshot_as_png
        driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[1]').click()
        driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[3]/div[2]/span[2]/div').click()
        question2 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[2]/div').screenshot_as_png
        driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[1]').click()
        try:
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[3]/div[2]/span[3]/div').click()
            question3 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[2]/div').screenshot_as_png
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[1]').click()
        except:
            print('2 bl')
        try:
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[3]/div[2]/span[4]/div').click()
            question4 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[2]/div').text
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[1]').click()
        except:
            bl = 0
        name = driver.find_element_by_xpath(
            f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/a[{student_number}]/div[2]/div').text
        print(question1)
        image = question1
        imageStream = io.BytesIO(image)
        doc = DocxTemplate("1.docx")
        sc1 = doc.new_subdoc()
        sc1.add_picture(imageStream, width=Inches(7))
        content = {'question1': sc1}
        doc.render(content)
        doc.save('test.docx')
        student_number += 1
    break
    driver.get(f'https://app.oqylyq.kz/t/assignments?page=2')
    sleep(1)




