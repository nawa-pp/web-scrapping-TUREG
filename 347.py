from pickle import TRUE
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By

# webdriver.Chrome('./chromedriver.exe').get('https://web.reg.tu.ac.th/registrar/class_info.asp?lang=th')
driver = webdriver.Chrome('./chromedriver.exe')  

driver.get('https://web.reg.tu.ac.th/registrar/class_info.asp?lang=th%27')
# find facultyid
driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[4]/td[2]/font[2]/select/option[11]").click()
# find semester
driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[6]/td[2]/table/tbody/tr[1]/td[2]/font[1]/select/option[2]").click()
# find acadyear
driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[6]/td[2]/table/tbody/tr[1]/td[2]/font[2]/select/option[7]").click()
# find CAMPUSID
driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[6]/td[2]/table/tbody/tr[2]/td[2]/select/option[1]").click()
# find LEVELID
driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[6]/td[2]/table/tbody/tr[3]/td[2]/select/option[2]").click()
# click! 
driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/table/tbody/tr[7]/td[2]/table/tbody/tr/td/font[3]/input").click()

data = pd.DataFrame()
for i in range(4,30):
    try:
        campus = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(i)+"]/td[2]/font").text
        courseCode = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(i)+"]/td[5]/font/a/b").text
        acadyear = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(i)+"]/td[3]/font").text
        courseName = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(i)+"]/td[6]/font").text
        unitOfCredit = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(i)+"]/td[7]/font").text
        section = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(i)+"]/td[8]/font/b").text
        time = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(i)+"]/td[9]/font").text
        examDateTime = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(i)+"]/td[10]/font").text
        seatquota = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(i)+"]/td[11]/font").text
        remain = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(i)+"]/td[12]/font").text
        status = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(i)+"]/td[13]/font").text 
        data = data.append(
            {
                'campus':campus,
                'courseCode':courseCode,
                'acadyear':acadyear,
                'courseName':courseName,
                'unitOfCredit':unitOfCredit,
                'section':section,
                'time':time,
                'examDateTime':examDateTime,
                'seatquota':seatquota,
                'remain' :remain,
                'status':status,
            }, ignore_index= TRUE
        )
    except:
        print("except")
        break

data.to_excel('CS347WTH.xlsx',sheet_name='Holy Sheet')
