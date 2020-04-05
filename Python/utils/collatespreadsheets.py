#!/usr/bin/env python
# coding: utf-8

# In[44]:

# written by : vishwanath kamath pethri
# Intent : Collate multiple spreadsheets with headers and multi sheets to one sheet

import pandas as pd
import csv
import os
get_ipython().system('pip install openpyxl')


# In[45]:


import openpyxl


# In[46]:



path = os.getcwd()
files = os.listdir(os.path.join(path, '<inputFolder>'))
files_xls = [f for f in files if f[-4:] == 'xlsx']

print(files_xls)


# In[54]:


with pd.ExcelWriter('<outputfile>.xlsx') as writer:
    for e in files_xls: 
        data = pd.read_excel(e, sheet_name=None, usecols = ['<generic_ColName1>','<generic_ColName2>','<generic_ColName3>','<generic_ColName3>'])
        df = pd.concat(data, ignore_index=True)
        df.to_excel(writer)


# In[ ]:




