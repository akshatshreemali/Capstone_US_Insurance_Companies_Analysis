# -*- coding: utf-8 -*-
"""
Created on Sat Jun  9 17:45:43 2018

@author: Akshat
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd

data = requests.get('http://www.state.nj.us/dobi/data/inscomp.htm').content
soup = BeautifulSoup(data, "html.parser")

table = soup.find(text="Company Name").find_parent("table")
columns=['Company Name','Address','City,State,Zip','Phone/NAIC']
lis=[] # creating empty list
for row in table.find_all("tr")[1:]:
    x=(([cell.get_text(strip=True) for cell in row.find_all("td")]))
    lis.append(x)
final_data=pd.DataFrame(lis)    
final_data.columns=columns
final_data.to_csv('C:/Users/Akshat/Desktop/Capstone_Data/NJ_Company.csv',index=False)



