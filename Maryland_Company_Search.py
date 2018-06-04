# -*- coding: utf-8 -*-
"""
Created on Sat Jun  2 23:41:06 2018

@author: Akshat
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import xlsxwriter
from pprint import pprint

workbook = xlsxwriter.Workbook(r'C:/Users/Akshat/Desktop/python/Scraped.xlsx')
worksheet = workbook.add_worksheet()
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

#   Write some data headers.
worksheet.write('A1', 'Company Name', bold)
worksheet.write('B1', 'NAIC#', bold)
worksheet.write('C1', 'Status', bold)
worksheet.write('D1', 'Business Street Address', bold)
worksheet.write('E1', 'Zip Code', bold)
worksheet.write('F1', 'Business Phone Number', bold)


#   Start from the first cell below the headers.
row = 1
col = 0

driver = webdriver.Chrome()
url='https://www.apps.insurance.maryland.gov/CompanyProducerInfo/InsCompanySearch.aspx?NAV=HOME'
driver.get(url)

time.sleep(30)



#['1','E','3','9','I','5','U','4','F','6','0','2']
select = Select(driver.find_element_by_id('text_status'))
dropdown=([o.text for o in select.options])
list_of_values_dropdown=dropdown[1:] # for testing let it start from 4
print(list_of_values_dropdown)
for i in list_of_values_dropdown:
    
    select = Select(driver.find_element_by_id('text_status'))
    select.select_by_visible_text(i)
    time.sleep(5)
    driver.find_element_by_id('maintable')
    driver.find_element_by_id('ContentPlaceHolder1_SearchButton').click()
    time.sleep(50)
    
    # change the entries to 100
    countHundred = driver.find_element_by_name("records_table_length")
    countHundred.click()
    countHundred.find_element(By.XPATH, '//option[text()="100"]').click()
    time.sleep(30)
    pages=driver.find_element_by_id('records_table_paginate')
    page=pages.find_element_by_tag_name('span')
    page_tags=page.find_elements_by_tag_name('a')
    final_page=page_tags[-1].text
    jam=int(final_page)
    xyz=0
    flag=0
    #driver.find_element_by_id("text_naic").click
    
    while xyz<jam:
            mainTable = driver.find_element_by_id("records_table")
            tableBody = mainTable.find_element_by_tag_name('tbody')
            tableRows = tableBody.find_elements_by_tag_name('tr')
            time.sleep(10)
            
            for singleRow in tableRows:
                tableData = singleRow.find_elements_by_tag_name('td')
            
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
            
                    row += 1
                  

            driver.find_element_by_id('records_table_next').click() 
            #time.sleep(10)
            xyz+=1              
    time.sleep(15)
               
workbook.close()