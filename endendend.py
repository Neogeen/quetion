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
driver.get(f'https://app.oqylyq.kz/t/assignments?page=1&query=%D0%9D%D0%B5%D1%84%D1%82%D0%B5%D1%85%D0%B8%D0%BC%D0%B8')
sleep(1)

for j in range(1,12):
    dl = 5
    if j == 11:
        nope = 0
        break
    sleep(2)
    exz = str(driver.find_element_by_xpath(f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[2]/div/a[{j}]/div[2]/div[1]/div[1]').text)
    c = driver.find_element_by_xpath(f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[2]/div/a[{j}]/div[1]/div/div/div/div[2]/div/div').text
    if c == '0':
        # driver.get(f'https://app.oqylyq.kz/t/assignments?page={i}')
        continue
    sleep(1)
    print(exz)


    driver.find_element_by_xpath(f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[2]/div/a[{j}]/div[2]/div[1]/div[1]').click()


    sleep(2)
    # driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/a').click()
    student_number = 1
    #d
    sleep(3)
    driver.find_element_by_xpath(
        '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[1]/div[2]/div/div/div[2]/div[2]').click()
    sleep(3)
    driver.find_element_by_xpath("//*[contains(text(), '100 студентов')]").click()
    for i in range(1, 130):
        isimage = 0
        bl = 2
        rp = 0
        pl = 0
        iskl = ''
        question3 = ""
        question4 = ""
        question5 = ""
        B4 = ''
        T5 = ''
        T5B1 = 0
        T5B2 = 0
        T5B3 = 0
        T5B4 = 0
        T1B5, T2B5, T3B5, T4B5, T5B5 = 0, 0, 0, 0, 0
        T4B1, T4B2, T4B3, T4B4, T4B5 = 0, 0, 0, 0, 0
        T1B1, T1B2, T1B3, T1B4, T1B5 = 0, 0, 0, 0, 0
        T1B4, T2B4, T3B4, T4B4, T5B4 = 0, 0, 0, 0, 0
        T2B1, T2B2, T2B3, T2B4, T2B5 = 0, 0, 0, 0, 0
        T3B1, T3B2, T3B3, T3B4, T3B5 = 0, 0, 0, 0, 0
        B4 = 0
        link = driver.current_url
        #Клик по студенту
        sleep(4)



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
        #------------------------------------------------------
        while True:
            try:
                name = driver.find_element_by_xpath(
                    f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/a[{student_number}]/div[2]/div').text
                break
            except:
                continue
        sleep(3)
        # БАЛЛЫ--------------------------------------------------------------------
        try:
            T1B1 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[4]/div[2]/div[1]/div/div[3]/div/span').text
        except:
            iskl += 'P '
            end = 0
        try:
            T1B2 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[4]/div[2]/div[2]/div/div[3]/div/span').text
        except:
            end = 0
        try:
            T1B3 = driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[4]/div[2]/div[3]/div/div[3]/div/span').text
        except:
            end = 0
        try:
            T2B1 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[5]/div[2]/div[1]/div/div[3]/div/span').text
            T2B2 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[5]/div[2]/div[2]/div/div[3]/div/span').text
            T2B3 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[5]/div[2]/div[3]/div/div[3]/div/span').text
        except:
            dl = 4
        try:
            T3B1 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[6]/div[2]/div[1]/div/div[3]/div/span').text
            T3B2 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[6]/div[2]/div[2]/div/div[3]/div/span').text
            T3B3 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[6]/div[2]/div[3]/div/div[3]/div/span').text
        except:
            dl = 4
        try:
            T4B1 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[7]/div[2]/div[1]/div/div[3]/div/span').text
            T4B2 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[7]/div[2]/div[2]/div/div[3]/div/span').text
            T4B3 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[7]/div[2]/div[3]/div/div[3]/div/span').text
        except:
            dl = 4
        try:
            T5B1 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[8]/div[2]/div[1]/div/div[3]/div/span').text
            T5B2 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[8]/div[2]/div[2]/div/div[3]/div/span').text
            T5B3 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[8]/div[2]/div[3]/div/div[3]/div/span').text
            T5B4 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[8]/div[2]/div[4]/div/div[3]/div/span').text
        except:
            dl = 4
        try:
            T1B4 = driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[4]/div[2]/div[4]/div/div[3]/div/span').text
        except:
            end = 0
        try:
            T2B4 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[5]/div[2]/div[4]/div/div[3]/div/span').text
        except:
            end = 0
        try:
            T3B4 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[6]/div[2]/div[4]/div/div[3]/div/span').text
        except:
            end = 0
        try:
            T4B4 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[7]/div[2]/div[4]/div/div[3]/div/span').text
        except:
            end = 0
        try:
            T5B4 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[8]/div[2]/div[4]/div/div[3]/div/span').text
        except:
            nope = 0
        sleep(2)
        try:
            T1 = driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[4]/div[1]').text
        except:
            student_number +=1
            continue
        try:
            T2 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[5]/div[1]').text
        except:
            end = 0
        try:
            T3 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[6]/div[1]').text
            T4 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[7]/div[1]').text
        except:
            end = 0
        try:
            T5 = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[8]/div[1]').text
        except:
            nope = 0
        try:
            T1B5 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[4]/div[2]/div[5]/div/div[3]/div').text
            print(T1B5)
        except:
            end = 0

        try:
            T2B5 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[5]/div[2]/div[5]/div/div[3]/div').text
        except:
            end = 0
        try:
            T3B5 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[6]/div[2]/div[5]/div/div[3]/div/span').text
        except:
            end = 0
        try:
            T4B5 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[7]/div[2]/div[5]/div/div[3]/div/span').text
        except:
            end = 0
        try:
            T5B5 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[8]/div[2]/div[5]/div/div[3]/div/span').text
        except:
            end = 0
        exz1 = exz.replace(':', '')
        # ------------------------------------------------------------------------------------
        while True:
            try:
                driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[3]/div[2]/div[1]/div/div[3]').click()
                #/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[3]/div[2]/span[1]/div
                break
            except:
                print('x')
                continue


        #sleep(1)
        try:
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div/div//p[img]')
            iskl += 'Вопрос 1 '
            isimage = 1
        except:
            end = 0
        sleep(2)
        question1 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[2]').text
        sleep(3)
        driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[1]').click()
        #Клик по второму вопросу
        try:
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[3]/div[2]/div[2]/div/div[3]/div').click()
        except:
            end = 0
        #sleep(1)
        try:
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div/div//p[img]')
            iskl += 'Вопрос 2 '
            isimage = 1
        except:
            end = 0
        sleep(2)
        question2 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[2]').text
        sleep(3)
        driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[1]').click()
        sleep(1)
        # 3 вопрос
        try:
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[3]/div[2]/div[3]/div/div[3]/div').click()
            #sleep(1)
            try:
                driver.find_element_by_xpath(
                    '/html/body/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div/div//p[img]')
                iskl += 'Вопрос 3 '
                isimage = 1
            except:
                end = 0
            sleep(2)
            question3 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[2]').text
            sleep(2)
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[1]').click()
        except:
            print('2 bl')
        try:
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[3]/div[2]/div[4]/div').click()
            #sleep(1)
            try:
                driver.find_element_by_xpath(
                    '/html/body/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div/div//p[img]')
                iskl += 'Вопрос 4'
                isimage = 1
            except:
                end = 0
            sleep(2)
            question4 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[2]').text
            sleep(2)
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[1]').click()
        except:
            bl = 0
        try:
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/div/div/div[5]/div[2]/div[3]/div[2]/div[5]/div/div[3]/div').click()
            sleep(1)
            try:
                driver.find_element_by_xpath(
                    '/html/body/div[1]/div/div/div[2]/div/div/div[2]/div/div[2]/div/div//p[img]')
                iskl += 'Вопрос 5'
            except:
                end = 0
            sleep(1)
            question5 = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[2]').text
            sleep(2)
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[1]').click()
        except:
            if bl == 0:
                end = 0
            else:
                bl = 1
            print(1)
        #--------------------------------------------------------------------------------------------------------
        sleep(1)

        #Сохранение в документ------------------------------------------------------------------------------------------
        doc = DocxTemplate(f"{bl}.docx")

        if isimage == 0:
            context = {'Fio': name, 'Exz': exz1, 'question1' : question1, 'question2' : question2, 'question3' : question3, 'question4' : question4,'question5':question5,'T1B5': T1B5,'T2B5': T2B5,'T3B5': T3B5,'T4B5': T4B5,'T5B5': T5B5, 'T1': T1, 'T2': T2, 'T3': T3, 'T4' : T4, 'T5': T5, 'T1B1': T1B1, 'T1B2': T1B2, 'T1B3' : T1B3, 'T1B4': T1B4, 'T2B1': T2B1, 'T2B2': T2B2, 'T2B3' : T2B3, 'T2B4': T2B4, 'T3B1': T3B1, 'T3B2': T3B2, 'T3B3' : T3B3, 'T3B4': T3B4, 'T4B1': T4B1, 'T4B2': T4B2, 'T4B3' : T4B3, 'T4B4': T4B4, 'T5B1': T5B1, 'T5B2': T5B2, 'T5B3' : T5B3, 'T5B4': T5B4}
        else:
            context = {'link': link, 'Fio': name, 'Exz': exz1, 'question1': question1,
                       'question2': question2, 'question3': question3, 'question4': question4,'question5':question5 , 'T1': T1, 'T2': T2,
                       'T3': T3, 'T4': T4, 'T5': T5, 'T1B1': T1B1, 'T1B2': T1B2, 'T1B3': T1B3, 'T1B4': T1B4,
                       'T2B1': T2B1, 'T2B2': T2B2, 'T2B3': T2B3, 'T2B4': T2B4, 'T3B1': T3B1, 'T3B2': T3B2, 'T3B3': T3B3,
                       'T3B4': T3B4, 'T4B1': T4B1, 'T4B2': T4B2, 'T4B3': T4B3, 'T4B4': T4B4, 'T5B1': T5B1, 'T5B2': T5B2,
                       'T5B3': T5B3, 'T5B4': T5B4,'T1B5': T1B5,'T2B5': T2B5,'T3B5': T3B5,'T4B5': T4B5,'T5B5': T5B5}
        doc.render(context)
        doc.save(f"{name} {exz1} ({iskl}).docx")
        try:
            driver.find_element_by_xpath(f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/a[{student_number}]/div[2]/div').click()
        except:
            sleep(3)
            driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div[1]').click()
            sleep(1)
            driver.find_element_by_xpath(
                f'/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/div[1]/div/div[8]/div/div[2]/div/a[{student_number}]/div[2]/div').click()
        student_number += 1
    driver.get(f'https://app.oqylyq.kz/t/assignments?page=23')
    sleep(1)




