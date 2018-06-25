# -*- coding: utf-8 -*-
"""
Created on Sun Jun 10 02:10:20 2018

@author: Akshat
"""

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

workbook = xlsxwriter.Workbook(r'C:/Users/Akshat/Desktop/Capstone_Data/Maryland_Company_Full.xlsx')
worksheet = workbook.add_worksheet()
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

#   Write some data headers.
worksheet.write('A1', 'Company Name', bold)
worksheet.write('B1', 'Previous Company Name', bold)
worksheet.write('C1', 'Merge Company Name', bold)
worksheet.write('D1', 'Address Line 1', bold)
worksheet.write('E1', 'Address Line 2', bold)
worksheet.write('F1', 'City,State,Zip', bold)
worksheet.write('G1', 'Company Website/ Email', bold)
worksheet.write('H1', 'NAIC', bold)
worksheet.write('I1', 'Phone Number', bold)
worksheet.write('J1', 'State of Domicile', bold)
worksheet.write('K1', 'Company Status', bold)
worksheet.write('L1', 'Application Field', bold)
worksheet.write('M1', 'Kinds of Insurance', bold)
worksheet.write('N1', 'Original Approval Date', bold)

#   Start from the first cell below the headers.
row = 0
col = 0

driver = webdriver.Chrome()
url='https://www.apps.insurance.maryland.gov/CompanyProducerInfo/InsCompanySearch.aspx?NAV=HOME'
driver.get(url)

time.sleep(50)



#['1','E','3','9','I','5','U','4','F','6','0','2']
select = Select(driver.find_element_by_id('text_status'))
dropdown=([o.text for o in select.options])
list_of_values_dropdown=dropdown[1:3] # for testing let it start from 4
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
    countHundred.find_element(By.XPATH, '//option[text()="10"]').click()
    time.sleep(10)
    pages=driver.find_element_by_id('records_table_paginate')
    page=pages.find_element_by_tag_name('span')
    page_tags=page.find_elements_by_tag_name('a')
    final_page=page_tags[-1].text
    jam=int(final_page)
    xyz=0
    flag=0
    #driver.find_element_by_id("text_naic").click
    
    while xyz<2:
            mainTable = driver.find_element_by_id("records_table")
            tableBody = mainTable.find_element_by_tag_name('tbody')
            tableRows = tableBody.find_elements_by_tag_name('tr')

            
            for singleRow in tableRows:
                window_before = driver.window_handles[0]
                singleRow.find_element_by_link_text('Information and Documents.').click()
                time.sleep(10)
                window_after = driver.window_handles[1]
                driver.switch_to_window(window_after)
                
                #another_window = list(set(driver.window_handles) - {driver.current_window_handle})[0]                
                mainTable2=driver.find_element(By.XPATH, '/html/body/form/table[1]')
                tableBody2 = mainTable2.find_element_by_tag_name('tbody')
                tableRows2 = tableBody2.find_elements_by_tag_name("tr")                
                
                i = 1
                col = 0
                row += 1
                for singleRow2 in tableRows2:
                    #time.sleep(5)
                    tableData2 = singleRow2.find_elements_by_tag_name("td")

                    time.sleep(3)
                    #zen=tableData2.find_element_by_tag_name('span')
                    #print(len(tableData2))
                    
                    for x in tableData2:
                        if len(tableData2) > 1:
                            #write data to excel sheet
                            if ((i % 2) == 0):
                                worksheet.write(row, col, x.text)
                                col += 1
                                #print('this is i',i)
                                #print('this is col',col)
                                #print(x.text)
                            i += 1
      
                elem=driver.find_element_by_tag_name("body")
                elem.send_keys(Keys.CONTROL + 'W')
                driver.close()
                driver.switch_to.window(window_before)     

            driver.find_element_by_id('records_table_next').click() 
            #time.sleep(10)
            xyz+=1 
             
    time.sleep(15)
               
workbook.close()