# -*- coding: utf-8 -*-
"""
Created on Fri Jun 29 01:53:28 2018

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
from selenium.common.exceptions import NoSuchElementException
import time
import xlsxwriter
from pprint import pprint

workbook = xlsxwriter.Workbook(r'C:/Users/Akshat/Desktop/Capstone_Data/NJ_Company_Agent_Expired_All.xlsx')
worksheet = workbook.add_worksheet()
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

#   Write some data headers.
worksheet.write('A1', 'Licensee Name', bold)
worksheet.write('B1', 'Ref Num', bold)
worksheet.write('C1', 'Business Address Name', bold)
worksheet.write('D1', 'License Type', bold)
worksheet.write('E1', 'Status', bold)
worksheet.write('G1', 'Authorities', bold)


#   Start from the first cell below the headers.
row = 1
col = 0

driver = webdriver.Firefox()
url='https://www20.state.nj.us/DOBI_LicSearch/insSearch.jsp?pager.offset=0&Refnum='
driver.get(url)

time.sleep(10)

#status = driver.find_element_by_id("text_status")
#status.click()
#status.find_element(By.XPATH, '//option[text()="1"]').click()
l=[]
select = Select(driver.find_element_by_name('LicenseType')) # this gets the dropdown list
dropdown=([o.text for o in select.options])
list_of_values_dropdown=dropdown[1:] # first character is blank so start from 1
print(list_of_values_dropdown)
lst = Select(driver.find_element_by_name('LicenseStatus')) # gets the license status
option=([o.text for o in lst.options])
list_licence_status=option[2:] # this value is subjective to the load, default is 2:
print(list_licence_status)

for d in list_of_values_dropdown:
    
    select = Select(driver.find_element_by_name('LicenseType'))# this is repeated to avoid 'Stale element error'
    select.select_by_visible_text(d) # selects the elements by text
    time.sleep(3)
   
    for j in list_licence_status:
        select = Select(driver.find_element_by_name('LicenseType'))
        select.select_by_visible_text(d)
        lst = Select(driver.find_element_by_name('LicenseStatus'))
        lst.select_by_visible_text(j)
        driver.find_element_by_css_selector("input[type='button']").click()
        time.sleep(60) # wait for the page to load
        url='https://www20.state.nj.us/DOBI_LicSearch/LicenseeSearchResults.jsp?pager.offset=128200'
		 #url='https://www20.state.nj.us/DOBI_LicSearch/LicenseeSearchResults.jsp?pager.offset=128200'
        driver.get(url) # url gets changed
		# click on search button
        #driver.find_element_by_xpath("//input[@type='button']//li[text()='Search']").click()
        #driver.find_element_by_xpath('ContentPlaceHolder1_SearchButton').click()
        time.sleep(60)
        # try catch loop checks if there are multiple pages or just one page. 
        # Depending on that, the loop will run
        try: 
            pages=driver.find_element_by_class_name('rnav')
            page_tags=pages.find_elements_by_tag_name('strong')
            final_page=(page_tags[-1].text)# total number of records
            jam=int(final_page)# converting into int
            total_iter=round(jam/50)
            print(total_iter)
        except :
            total_iter=0
        
        xyz=0
        while xyz<=4:
                time.sleep(10)
                mainTable=driver.find_element(By.XPATH, '/html/body/div/div/form[2]/center/table[1]')
                tableBody = mainTable.find_element_by_tag_name('tbody')
                tableRows = tableBody.find_elements_by_tag_name('tr')
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
                        worksheet.write(row, col + 5,  tableData[i].text)
                        row += 1
                        
                             
                try:
                    if xyz==0:
                         driver.find_element_by_xpath('/html/body/div/div/form[2]/center/div[2]/a[10]').click() 
                    else :
                        driver.find_element_by_xpath('/html/body/div/div/form[2]/center/div[2]/a[11]').click()
                except:
                    print('Next not required')
                #driver.find_element_by_xpath("/html/body/div/form[1]/center/div[1]/")
                #time.sleep(10)
                xyz+=1            
        driver.find_element_by_xpath('//*[@id="main"]/form[2]/center/table[2]/tbody/tr/td/font/input').click()               
    time.sleep(15)
               
workbook.close()



    

