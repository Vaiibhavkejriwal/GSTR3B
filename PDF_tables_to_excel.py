#!/usr/bin/env python
# coding: utf-8

# In[187]:


import tabula
import pandas as pd
from pandas import DataFrame
import numpy
import os
import glob
import time

def pdf_table(filepath):
    file = filepath
    pth=os.path.dirname(file)
    files = glob.glob(os.path.join(pth, '*.pdf*'))

    for f in files:
        df = tabula.read_pdf(f, pages = 'all')
        pth=os.path.dirname(f)
        new_extension = '.xlsx'
        filename = os.path.splitext(f)[0]
        newfile= os.path.join(pth,filename + new_extension)
        writer = pd.ExcelWriter(newfile, engine='xlsxwriter')
        try:
            for i in list(range(0, 100)):
                d = pd.DataFrame(df[i + 2])
                d.to_excel(writer, sheet_name = 'Table' + str(i + 1), index = False)
        except IndexError:
            writer.save()