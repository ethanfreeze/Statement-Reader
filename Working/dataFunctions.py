#!/usr/bin/env python
# coding: utf-8

# In[11]:


import re
import pandas as pd
import PyPDF2
import os
from openpyxl import load_workbook, Workbook
from IPython.display import display


# In[12]:


def getData():
    lineData = {'Date': [],
                'Transaction': [],
                'Amount': []}
    for x in os.listdir():
        if x.endswith('.pdf'):
            path = os.getcwd() + '\\' + x
            file = open(path, 'rb')
            readFile = PyPDF2.PdfFileReader(file)


            totalPages = readFile.numPages

            for i in range(totalPages):

                pageObj = readFile.getPage(i)
                pageText = pageObj.extractText

                newTrans = re.compile(r'[A-Z][a-z]{2} \d{2}\s')
                moneyRe = re.compile(r'\d{1,}\.\d{2}')

                for line in pageText(pageObj).split('\n'):
                    line = re.sub(r',','',line)
                    line = re.sub(r'\$\s','',line)
                    if newTrans.match(line):

                        newValue = re.split(newTrans, line)
                        newValue = ' '.join(newValue)
                        newValue = re.split(moneyRe, newValue)
                        newValue = ' '.join(newValue)

                        newKey = newTrans.findall(line)
                        newKey = ' '.join(newKey)

                        newAmount = moneyRe.findall(line)
                        newAmount = ' '.join(newAmount)
                        

                        lineData['Date']+=[newKey]
                        lineData['Transaction']+=[newValue]
                        lineData['Amount']+=[newAmount]
        dataFrame = pd.DataFrame(lineData)
        dataFrame.to_csv('Transaction Data.csv')
    return dataFrame


# In[13]:


def categorize():
    path = os.getcwd() + '\\' + 'Bank Data.xlsx'
    
    dataFrame = pd.DataFrame(getData())
    
    with pd.ExcelWriter('Bank Data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        dataFrame.to_excel(writer, sheet_name='Main')
    
    category = input('What Keyword Would You Like to categorize? ')
    newCategory = dataFrame.query('Transaction.str.contains(@category)', engine='python')
    
    category = category.lower()
    newCategory = newCategory.append(dataFrame.query('Transaction.str.contains(@category)', engine='python'))
    
    category = category.upper()
    newCategory = newCategory.append(dataFrame.query('Transaction.str.contains(@category)', engine='python'))
    
    category = category.capitalize()
    newCategory = newCategory.append(dataFrame.query('Transaction.str.contains(@category)', engine='python'))
    display(newCategory)
    
    userIn = input(f'Would you like to categorize these transactions? (Y/N) ')
    userIn = userIn.lower()
    if userIn == "yes" or userIn == "y" or userIn == "yy" :
        userIn = input(f'Enter name of category: ')
        userIn=str(userIn)
        with pd.ExcelWriter('Bank Data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            newCategory.to_excel(writer, sheet_name=userIn)
        print('Done!')
    else:
        return


# In[25]:


def loadCategories():
    sheets_dict = pd.read_excel('Bank Data.xlsx', sheet_name=None)
    all_sheets = []
    
    for name, sheet in sheets_dict.items():
        sheet['sheet'] = name
        sheet = sheet.rename(columns=lambda x: x.split('\n')[-1])
        all_sheets.append(sheet)
    
    full_table = pd.concat(all_sheets)
    full_table.reset_index(inplace=True, drop=True)

    print(full_table)
    
    return full_table
loadCategories()
categorize()

# In[39]:




# In[ ]:





# In[ ]:




