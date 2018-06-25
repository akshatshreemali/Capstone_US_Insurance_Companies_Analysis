# -*- coding: utf-8 -*-
"""
Created on Sat Jun 16 23:50:30 2018

@author: Akshat
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import time
import xlsxwriter
from pprint import pprint

workbook = xlsxwriter.Workbook(r'C:/Users/Akshat/Desktop/Capstone_Data/PA_Licensee.xlsx')
worksheet = workbook.add_worksheet()
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

#   Write some data headers.
worksheet.write('A1', 'Last Name', bold)
worksheet.write('B1', 'First Name', bold)
worksheet.write('C1', 'NMiddle Name', bold)
worksheet.write('D1', 'Suffix', bold)
worksheet.write('E1', 'License Type', bold)
worksheet.write('F1', 'City', bold)
worksheet.write('G1', 'State', bold)
worksheet.write('H1', 'License Number', bold)

#   Start from the first cell below the headers.
row = 1
col = 0

driver = webdriver.Chrome()
url='http://apps02.ins.state.pa.us/producer/ilist1.asp'
driver.get(url)

time.sleep(10)

#status = driver.find_element_by_id("text_status")
#status.click()
#status.find_element(By.XPATH, '//option[text()="1"]').click()
cities=["Aliquippa","Allentown","Altoona","Arnold","Beaver Falls","Bethlehem","Bradford","Butler","Carbondale",
"Chester","Clairton","Coatesville","Connellsville","Corry","DuBois","Duquesne","Easton","Erie","Farrell","Franklin",
"Greensburg","Harrisburg","Hazleton","Hermitage","Jeannette","Johnstown","Lancaster","Latrobe","Lebanon","Lock Haven",
"Lower Burrell","McKeesport","Meadville","Monessen","Monongahela","Nanticoke","New Castle","New Kensington","Oil City","Parker",
"Philadelphia","Pittsburgh","Pittston","Pottsville","Reading","St. Marys","Scranton","Shamokin","Sharon","Sunbury","Titusville",
"Uniontown","Warren","Washington","Wilkes-Barre","Williamsport","York"]

#y=['Arnold']
for city in cities:
    inp=driver.find_element_by_xpath('//*[@id="wrap"]/div[2]/table/tbody/tr[2]/td/form[2]/table/tbody/tr[6]/td[2]/font/input')# getting city column
    inp.click()
    time.sleep(2)
    inp.send_keys(city)# filling the value with city
    time.sleep(5)
    state=driver.find_element_by_xpath('//*[@id="wrap"]/div[2]/table/tbody/tr[2]/td/form[2]/table/tbody/tr[7]/td[2]/font/select')
    state.click()
    state.find_element(By.XPATH, '//option[text()="Pennsylvania"]').click()
    form=driver.find_element_by_xpath('//*[@id="wrap"]/div[2]/table/tbody/tr[2]/td/form[2]/table')# getting the table for submit button
    form.find_element_by_name('btnSubmit').click() # click on submit button
    time.sleep(15)
    # below try catch will help in running the code in case there's no data
    try:
        alpha=driver.find_element_by_xpath('//*[@id="wrap"]/div[2]/table/tbody/tr[2]/td/p[2]').text
        alpha=alpha.split(': ')
        alpha=alpha[-1]
        print('Total data in this city is', alpha)
    except:
        driver.find_element_by_name('btnAnother').click()
    beta=2 # this variable will loop to get data for license number
    y='//*[@id="wrap"]/div[2]/table/tbody/tr[2]/td/table/tbody/tr['
    mainTable = driver.find_element_by_xpath('//*[@id="wrap"]/div[2]/table/tbody/tr[2]/td/table') # getting main table for data
    tableBody = mainTable.find_element_by_tag_name('tbody')
    tableRows = tableBody.find_elements_by_tag_name('tr')
    for singleRow in tableRows[1:]: # skipping first element since it contains column names
        tableData = singleRow.find_elements_by_tag_name('td')
        link=y+str(beta)+']/td[8]/input' # this concatenates sting required to fetch xpath dynamically
        licnum=singleRow.find_element_by_xpath(link).get_attribute('value') # gets license number for each table
        for x in range(1):     
        			i = 0
        			#write data to excel sheet
        			worksheet.write(row, col,     tableData[i].text)
        			i += 1
        			worksheet.write(row, col + 1,  tableData[i].text)
        			i += 1
        			worksheet.write(row, col + 2,  tableData[i].text)
        			i += 1
        			worksheet.write(row, col + 3,  tableData[i].text)
        			i += 1
        			worksheet.write(row, col + 4,  tableData[i].text)
        			i += 1
        			worksheet.write(row, col + 5,  tableData[i].text)
        			i += 1
        			worksheet.write(row, col + 6,  tableData[i].text)
        			i += 1
        			worksheet.write(row, col + 7,  licnum)
        			row += 1
        beta+=1           
           
    driver.find_element_by_xpath('//*[@id="wrap"]/div[2]/table/tbody/tr[2]/td/p[1]/a').click() # new search
    
workbook.close()      
          



    

