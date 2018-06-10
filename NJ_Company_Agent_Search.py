
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
from selenium.common.exceptions import NoSuchElementException
import time
import xlsxwriter
from pprint import pprint

workbook = xlsxwriter.Workbook(r'C:/Users/Akshat/Desktop/Capstone_Data/NJ_Company_Agent.xlsx')
worksheet = workbook.add_worksheet()
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

#   Write some data headers.
worksheet.write('A1', 'Licensee Name', bold)
worksheet.write('B1', 'Ref Num', bold)
worksheet.write('C1', 'Business Address Name', bold)
worksheet.write('D1', 'License Type', bold)
worksheet.write('E1', 'Status', bold)
worksheet.write('F1', 'Status URL', bold)
worksheet.write('G1', 'Authorities', bold)


#   Start from the first cell below the headers.
row = 1
col = 0

driver = webdriver.Chrome()
url='https://www20.state.nj.us/DOBI_LicSearch/insSearch.jsp?pager.offset=0&Refnum='
driver.get(url)

time.sleep(10)

#status = driver.find_element_by_id("text_status")
#status.click()
#status.find_element(By.XPATH, '//option[text()="1"]').click()
l=[]
select = Select(driver.find_element_by_name('LicenseType'))
dropdown=([o.text for o in select.options])
list_of_values_dropdown=dropdown[2:] # for testing let it start from 4
print(list_of_values_dropdown)
for i in list_of_values_dropdown:
    
    select = Select(driver.find_element_by_name('LicenseType'))
    select.select_by_visible_text(i)
    time.sleep(5)
    driver.find_element_by_css_selector("input[type='button']").click()
    #driver.find_element_by_xpath("//input[@type='button']//li[text()='Search']").click()
    #driver.find_element_by_xpath('ContentPlaceHolder1_SearchButton').click()
    time.sleep(50)
    try:
        pages=driver.find_element_by_class_name('rnav')
        page_tags=pages.find_elements_by_tag_name('strong')
        final_page=(page_tags[-1].text)# total number of records
        jam=int(final_page)# converting into int
        total_iter=round(jam/50)
    except :
        total_iter=0
    
    xyz=0
    while xyz<=1:
        
            mainTable=driver.find_element(By.XPATH, '//*[@id="main"]/form[2]/center/table[1]')
            tableBody = mainTable.find_element_by_tag_name('tbody')
            tableRows = tableBody.find_elements_by_tag_name('tr')
            time.sleep(10)
            
            for singleRow in tableRows[1:]:
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
                    worksheet.write(row, col + 5,  tableData[i].get_attribute('href'))
                    i += 1
                    worksheet.write(row, col + 6,  tableData[i].text)
                    row += 1
            try:
                if xyz==0:
                     driver.find_element_by_xpath('//*[@id="main"]/form[2]/center/div[2]/a[10]').click() 
                else :
                    driver.find_element_by_xpath('//*[@id="main"]/form[2]/center/div[2]/a[11]').click()
            except:
                print('Next not required')
            #driver.find_element_by_xpath("/html/body/div/form[1]/center/div[1]/")
            #time.sleep(10)
            xyz+=1            
    driver.find_element_by_xpath('//*[@id="main"]/form[2]/center/table[2]/tbody/tr/td/font/input').click()               
    time.sleep(15)
               
workbook.close()



    

