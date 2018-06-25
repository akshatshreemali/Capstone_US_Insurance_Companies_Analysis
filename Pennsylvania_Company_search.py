# -*- coding: utf-8 -*-
"""
Created on Fri Jun 15 21:48:55 2018

@author: Akshat
"""



from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
import time
import xlsxwriter
from pprint import pprint
import pandas as pd

workbook = xlsxwriter.Workbook(r'C:/Users/Akshat/Desktop/Capstone_Data/Pennsylvania3.xlsx')
worksheet = workbook.add_worksheet()
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

#   Write some data headers.
worksheet.write('A1', 'Company Name', bold)
worksheet.write('B1', 'NAIC', bold)
worksheet.write('C1', 'Mailing Address', bold)
worksheet.write('D1', 'Domicile', bold)
worksheet.write('E1', 'Phone', bold)
worksheet.write('F1', 'Type', bold)
worksheet.write('G1', 'Powers', bold)
worksheet.write('H1', 'Risk', bold)

#   Start from the first cell below the headers.
row = 0
col = 0

driver = webdriver.Chrome()
url='http://www.insurance.state.pa.us/dsf/gfsearch.html'
driver.get(url)

time.sleep(10)
  
# //*[@id="FormsRadioButton01"]
xpath='FormsRadioButton'
lis=['01','02','03','04',"05","06","07","08","09","10","11","12","13","14","15","16","17","18","19","20","21","22"]
#test=['01','02','03']
for i in lis:
    x=xpath+i
    a=("'{}'".format(x))
    print('this is a',a)
    rowsda=driver.find_element_by_id(x)
    rowsda.click()
    driver.find_element_by_xpath('//*[@id="FormsButton4"]').click()
    
    mainTable=driver.find_element_by_xpath('//*[@id="wrap"]/div[2]/table')
    tableBody = mainTable.find_element_by_tag_name('tbody')
    tableRows = tableBody.find_elements_by_tag_name('tr')
    alpha=driver.find_element_by_xpath('//*[@id="wrap"]/div[2]/p[1]').text
    alpha=alpha.split(': ')
    alpha=alpha[-1]
    alpha=int(alpha)
    y='//*[@id="wrap"]/div[2]/table/tbody/tr['
    for j in range(2,alpha):
        
                link=y+str(j)+']/td[3]/a'
                c=driver.find_element_by_xpath(link).text
                #print('this is c',c)
                driver.find_element_by_xpath(link).click()
                #beta[2].click()
                mainTable2=driver.find_element_by_xpath('//*[@id="wrap"]/div[2]')
                #tableBody2 = mainTable2.find_element_by_tag_name('tbody')
                tableRows2 = mainTable2.find_elements_by_tag_name("p")                
                
                i = 1
                col = 0
                row += 1
                    
                for x in tableRows2[1:]:
                    #if len(tableData2) > 1:
                        #write data to excel sheet
                        #if ((i % 2) == 0):
                            worksheet.write(row, col, c)
                            col+=1
                            worksheet.write(row, col, x.text)
                            col += 1
                            #print('this is i',i)
                            #print('this is col',col)
                            #print(x.text)
                        #i += 1
                time.sleep(5)        
                driver.find_element_by_xpath('//*[@id="FormsButton4"]').click()
                time.sleep(10) 
                j+=1
    driver.back()
    time.sleep(6)        
        
workbook.close()
time.sleep(5)

driver.close

def data_clean():

    
    df=pd.read_excel('C:/Users/Akshat/Desktop/Capstone_Data/Pennsylvania3.xlsx')
    df.drop(['NAIC','Mailing Address','Phone','Powers','Unnamed: 8', 'Unnamed: 10'
           , 'Unnamed: 12', 'Unnamed: 14',
           'Unnamed: 15', 'Unnamed: 16','Unnamed: 18','Unnamed: 20', 'Unnamed: 21', 'Unnamed: 22',
           'Unnamed: 23', 'Unnamed: 24','Unnamed: 26',
           'Unnamed: 27', 'Unnamed: 28', 'Unnamed: 29', 'Unnamed: 30',
           'Unnamed: 31', 'Unnamed: 32', 'Unnamed: 33', 'Unnamed: 34'],axis=1,inplace=True)
    df1 = pd.DataFrame(df.Domicile.str.split(':',1).tolist(),
                                       columns = ['junk','NAIC'])
    df2 = pd.DataFrame(df.Type.str.split('\n',1).tolist(),
                                       columns = ['junk2','Home Address'])
    df3 = pd.DataFrame(df.Risk.str.split('\n',1).tolist(),
                                       columns = ['junk3','Mailing Address'])
    df4 = pd.DataFrame(df['Unnamed: 9'].str.split(':',1).tolist(),
                                       columns = ['junk4','Domicile'])
    df5 = pd.DataFrame(df['Unnamed: 11'].str.split(':',1).tolist(),
                                       columns = ['junk5','Phone'])
    df6 = pd.DataFrame(df['Unnamed: 13'].str.split(':',1).tolist(),
                                       columns = ['junk6','Type'])
    df7 = pd.DataFrame(df['Unnamed: 17'].str.split('\n',1).tolist(),
                                       columns = ['junk7','Powers'])
    
    df=df.merge(df1,left_index=True,right_index=True)
    df=df.merge(df2,left_index=True,right_index=True)
    df=df.merge(df3,left_index=True,right_index=True)
    df=df.merge(df4,left_index=True,right_index=True)
    df=df.merge(df5,left_index=True,right_index=True)
    df=df.merge(df6,left_index=True,right_index=True)
    df=df.merge(df7,left_index=True,right_index=True)
    
    df.drop(['Domicile_x', 'Type_x', 'Risk','Unnamed: 9',
           'Unnamed: 11', 'Unnamed: 13', 'Unnamed: 17','junk','junk2','junk3','junk4','junk5','junk6','junk7'],axis=1,inplace=True)
    
    df.drop_duplicates(subset=['Company Name', 'Unnamed: 19', 'Unnamed: 25', 'Unnamed: 35', 'NAIC',
           'Home Address', 'Mailing Address', 'Domicile_y', 'Phone', 'Type_y',
           'Powers'],inplace=True)
    
    columns=['Company Name','Risk','Information','Information_Date','NAIC','Home Address','Mailing Address','Domicile','Phone Number','Type','Powers']
    df.columns=columns
    
    df.to_excel('C:/Users/Akshat/Desktop/Capstone_Data/Pennsylvania3.xlsx',index=False)
    
    return print('data is cleaned and duplicates have been removed')
    
data_clean()
