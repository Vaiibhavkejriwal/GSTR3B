import tabula
import pandas as pd
from pandas import DataFrame
import numpy
import os
import glob
import datetime
from PDF_tables_to_excel import pdf_table

print("Change the names of the file to there respective months before running the program.")
file= input('Enter the File path of the PDFs: ')
            
t1 = datetime.datetime.now()
pdf_table(file)
pth=os.path.dirname(file)

def isfloat(value):
  try:
    float(value)
    return True
  except ValueError:
    return False

newfile= os.path.join(pth,'GSTRECO.xlsx')
writer = pd.ExcelWriter(newfile, engine='xlsxwriter')
files = glob.glob(os.path.join(pth, '*.xls*'))
df3 = pd.DataFrame()
df2 = pd.DataFrame()
df4 = pd.DataFrame()
df5 = pd.DataFrame()
df6 = pd.DataFrame()
df7 = pd.DataFrame()
df8 = pd.DataFrame()
df9 = pd.DataFrame()
df10 = pd.DataFrame()
df11 = pd.DataFrame()
df12 = pd.DataFrame()
df13 = pd.DataFrame()
df14 = pd.DataFrame()
df15 = pd.DataFrame()
df16 = pd.DataFrame()
df17 = pd.DataFrame()
df18 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f)
    u = name
    p =  df.iloc[1,1]
    q = df.iloc[1,2]
    r = df.iloc[1,3]
    s = df.iloc[1,4]
    t = df.iloc[1,5]
    if isfloat(p) == True:
        p = float(p)
    else:
        p = 0

    if isfloat(q) == True:
        q = float(q)
    else:
        q = 0
    if isfloat(q) == True:
        r = float(r)
    else:
        r = 0
    if isfloat(s) == True:
        s = float(s)
    else:
        s = 0

    if isfloat(t) == True:
        t = float(t)
    else:
        t = 0
    v = int(q) + int(r) + int(s) + int(t)
    row = [u, p, q, r, s, t, v]
    df2 = pd.DataFrame(row).T
    df3 = df3.append(df2)
df3.columns = ['Months', 'Taxable Value', 'IGST', 'CGST', 'SGST', 'Cess', 'Total Tax Liability']
df3.to_excel(writer, sheet_name = 'Outward other than nil rated', index = False)


# In[20]:


df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f)
    u = name
    p = df.iloc[3,1]
    q = df.iloc[3,2]
    r = df.iloc[3,3]
    s = df.iloc[3,4]
    t = df.iloc[3,5]
    if isfloat(p) == True:
        p = float(p)
    else:
        p = 0

    if isfloat(q) == True:
        q = float(q)
    else:
        q = 0
    if isfloat(r) == True:
        r = float(r)
    else:
        r = 0
    if isfloat(s) == True:
        s = float(s)
    else:
        s = 0

    if isfloat(t) == True:
        t = float(t)
    else:
        t = 0
    v = int(q) + int(r) + int(s) + int(t)
    row = [u, p, q, r, s, t, v]
    df2 = pd.DataFrame(row).T
    df4 = df4.append(df2)

df4.columns = ['Months', 'Taxable Value', 'IGST', 'CGST', 'SGST', 'Cess', 'Total Tax Liability']
df4.to_excel(writer, sheet_name = 'zero rated', index = False )


df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f)
    u = name
    p =  df.iloc[4,1]
    q = df.iloc[4,2]
    r = df.iloc[4,3]
    s = df.iloc[4,4]
    t = df.iloc[4,5]
    if isfloat(p) == True:
        p = float(p)
    else:
        p = 0

    if isfloat(q) == True:
        q = float(q)
    else:
        q = 0
    if isfloat(r) == True:
        r = float(r)
    else:
        r = 0
    if isfloat(s) == True:
        s = float(s)
    else:
        s = 0

    if isfloat(t) == True:
        t = float(t)
    else:
        t = 0
    v = int(q) + int(r) + int(s) + int(t)
    row = [u, p, q, r, s, t, v]
    df2 = pd.DataFrame(row).T
    df5 = df5.append(df2)
df5.columns = ['Months', 'Taxable Value', 'IGST', 'CGST', 'SGST', 'Cess', 'Total Tax Liability']
df5.to_excel(writer, sheet_name = 'nil rated', index = False)

df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f)
    u = name
    p =  df.iloc[5,1]
    q = df.iloc[5,2]
    r = df.iloc[5,3]
    s = df.iloc[5,4]
    t = df.iloc[5,5]
    if isfloat(r) == True:
        p = float(p)
    else:
        p = 0

    if isfloat(q) == True:
        q = float(q)
    else:
        q = 0
    if isfloat(q) == True:
        r = float(r)
    else:
        r = 0
    if isfloat(s) == True:
        s = float(s)
    else:
        s = 0

    if isfloat(t) == True:
        t = float(t)
    else:
        t = 0
    v = int(q) + int(r) + int(s) + int(t)
    row = [u, p, q, r, s, t, v]
    df2 = pd.DataFrame(row).T
    df6 = df6.append(df2)
df6.columns = ['Months', 'Taxable Value', 'IGST', 'CGST', 'SGST', 'Cess', 'Total Tax Liability']
df6.to_excel(writer, sheet_name = 'reverse charge', index = False)


df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f)
    u = name
    p =  df.iloc[6,1]
    q = df.iloc[6,2]
    r = df.iloc[6,3]
    s = df.iloc[6,4]
    t = df.iloc[6,5]
    if isfloat(p) == True:
        p = float(p)
    else:
        p = 0

    if isfloat(q) == True:
        q = float(q)
    else:
        q = 0
    if isfloat(r) == True:
        r = float(r)
    else:
        r = 0
    if isfloat(s) == True:
        s = float(s)
    else:
        s = 0

    if isfloat(t) == True:
        t = float(t)
    else:
        t = 0
    v = int(q) + int(r) + int(s) + int(t)
    row = [u, p, q, r, s, t, v]
    df2 = pd.DataFrame(row).T
    df7 = df7.append(df2)
df7.columns = ['Months', 'Taxable Value', 'IGST', 'CGST', 'SGST', 'Cess', 'Total Tax LIability']
df7.to_excel(writer, sheet_name = 'Non Gst Outward', index = False)


df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f, sheet_name = 1)
    u = name
    p =  df.iloc[1,1]
    q = df.iloc[1, 2]
    row = [u, p, q]
    df2 = pd.DataFrame(row).T
    df8 = df8.append(df2)
df8.columns = ['Months', 'Taxable Value', 'IGST']
df8.to_excel(writer, sheet_name = 'unregistered person', index = False)


df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f, sheet_name = 1)
    u = name
    p =  df.iloc[2,1]
    q = df.iloc[2,2]
    row = [u, p, q]
    df2 = pd.DataFrame(row).T
    df9 = df9.append(df2)
df9.columns = ['Months', 'Taxable Value', 'IGST']
df9.to_excel(writer, sheet_name = 'compostion', index = False)


# In[167]:


df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f, sheet_name = 1)
    u = name
    p =  df.iloc[2,1]
    q = df.iloc[2,2]
    row = [u, p, q]
    df2 = pd.DataFrame(row).T
    df10 = df10.append(df2)
df10.columns = ['Months', 'Taxable Value', 'IGST']
df10.to_excel(writer, sheet_name = 'UIN', index = False)


# In[168]:


df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f, sheet_name = 2)
    u = name
    p =  df.iloc[0,1]
    q = df.iloc[0,2]
    r = df.iloc[0,3]
    s = df.iloc[0,4]
    row = [u, p, q, r, s]
    df2 = pd.DataFrame(row).T
    df11 = df11.append(df2)
df11.columns = ['Months', 'IGST', 'CGST', 'SGST', 'Cess']
df11.to_excel(writer, sheet_name = 'ITC available', index = False)

df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f, sheet_name = 2)
    u = name
    p =  df.iloc[1,1]
    q = df.iloc[1,2]
    r = df.iloc[1,3]
    s = df.iloc[1,4]
    row = [u, p, q, r, s]
    df2 = pd.DataFrame(row).T
    df12 = df12.append(df2)
df12.columns = ['Months', 'IGST', 'CGST', 'SGST', 'Cess']
df12.to_excel(writer, sheet_name = 'ITC reversed', index = False)

df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f, sheet_name = 2)
    u = name
    p =  df.iloc[2,1]
    q = df.iloc[2,2]
    r = df.iloc[2,3]
    s = df.iloc[2,4]
    row = [u, p, q, r, s]
    df2 = pd.DataFrame(row).T
    df13 = df13.append(df2)
df13.columns = ['Months', 'IGST', 'CGST', 'SGST', 'Cess']
df13.to_excel(writer, sheet_name = 'Net ITC', index = False)


df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f, sheet_name = 2)
    u = name
    p =  df.iloc[3,1]
    q = df.iloc[3,2]
    r = df.iloc[3,3]
    s = df.iloc[3,4]
    row = [u, p, q, r, s]
    df2 = pd.DataFrame(row).T
    df14 = df14.append(df2)
df14.columns = ['Months', 'IGST', 'CGST', 'SGST', 'Cess']
df14.to_excel(writer, sheet_name = 'Ineligible ITC', index = False)

df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f, sheet_name = 3)
    u = name
    p =  df.iloc[0,1]
    q = df.iloc[0,2]
    row = [u, p, q]
    df2 = pd.DataFrame(row).T
    df15 = df15.append(df2)
df15.columns = ['Months', 'Inter state supplies', 'Intra state supplies']
df15.to_excel(writer, sheet_name = 'from composition', index = False)

df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f, sheet_name = 1)
    u = name
    p =  df.iloc[1,1]
    q = df.iloc[1,2]
    row = [u, p, q]
    df2 = pd.DataFrame(row).T
    df16 = df16.append(df2)
df16.columns = ['Months', 'Inter state supplies', 'Intra state supplies']
df16.to_excel(writer, sheet_name = 'Non GST', index = False)

df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f, sheet_name = 4)
    u = name
    p =  df.iloc[0,1]
    q = df.iloc[0,2]
    r = df.iloc[0,3]
    s = df.iloc[0,4]
    if isfloat(p) == True:
        p = float(p)
    else:
        p = 0

    if isfloat(q) == True:
        q = float(q)
    else:
        q = 0
    if isfloat(q) == True:
        r = float(r)
    else:
        r = 0
    if isfloat(s) == True:
        s = float(s)
    else:
        s = 0
    row = [u, p, q, r, s]
    df2 = pd.DataFrame(row).T
    df17 = df17.append(df2)
df17.columns = ['Months', 'IGST', 'CGST', 'SGST', 'Cess']
df17.to_excel(writer, sheet_name = 'Interest', index = False)

df2 = pd.DataFrame()
for f in files:
    base= os.path.basename(f)
    name = os.path.splitext(base)[0]
    df = pd.read_excel(f, sheet_name = 4)
    u = name
    p =  df.iloc[1,1]
    q = df.iloc[1,2]
    r = df.iloc[1,3]
    s = df.iloc[1,4]
    row = [u, p, q, r, s]
    df2 = pd.DataFrame(row).T
    df18 = df18.append(df2)

df18.columns = ['Months', 'IGST', 'CGST', 'SGST', 'Cess']
df18.to_excel(writer, sheet_name = 'Late fees', index = False)

df200 = pd.concat([df4, df5, df6, df7, df8, df9, df10, df11, df12, df13, df14, df15, df16, df17, df18], axis = 1)
df200 = df200.drop(['Months'], axis = 1)
df200 = pd.concat([df3, df200], axis = 1)
df200.to_excel(writer, sheet_name = 'GSTR3B', index = False)

writer.save()
t2 = datetime.datetime.now()
t = t2 - t1
t = int(t.total_seconds()*1)
print(f'Complete GSTR 3B has been generated in excel and it took only {t} seconds to do it. Cheers to saving a lot of time!')
print('Thank you for running the code :)) ')

