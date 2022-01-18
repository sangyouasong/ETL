#install python-docx library for extracting data from docx documents
!pip install python-docx

#import libraries 
from docx.api import Document
from docx import Document 
import pandas as pd

#find the path of targeted documents for table extraction 
!ls Desktop/*.docx

India = Document ('Desktop/2017 India FSSA.docx')

India.tables

for i in range(len(India.tables)):
    table = India.tables[i]
    data1= []
    keys = None
    row_data1 = None
    for j, row in enumerate(table.rows):
        text = (cell.text for cell in row.cells)
        if j == 0:
            keys = tuple(text)
            continue
        row_data1= dict(zip(keys,text))
        data1.append(row_data1)
    df1 = pd.DataFrame(data1)



print(df1)

df1.to_excel(r'Desktop/tables/India FSSA.xlsx', index = False, header = True)

##Extracting table from 2017 New Zealand 

NZ = Document ('Desktop/2017 New Zealand FSSA.docx')

for i in range(len(NZ.tables)):
    table = NZ.tables[i]
    data2= []
    keys = None
    row_data2 = None
    for j, row in enumerate(table.rows):
        text = (cell.text for cell in row.cells)
        if j == 0:
            keys = tuple(text)
            continue
        row_data2= dict(zip(keys,text))
        data2.append(row_data2)
    df2 = pd.DataFrame(data2)



df2.to_excel(r'Desktop/tables/NZ FSSA.xlsx', index = False, header = True)

Jamaica = Document ('Desktop/2018 Jamaica FSSA.docx')


for i in range(len(Jamaica.tables)):
    table = Jamaica.tables[i]
    data3= []
    keys = None
    row_data3 = None
    for j, row in enumerate(table.rows):
        text = (cell.text for cell in row.cells)
        if j == 0:
            keys = tuple(text)
            continue
        row_data3= dict(zip(keys,text))
        data3.append(row_data3)
    df3 = pd.DataFrame(data3)



df3.to_excel(r'Desktop/tables/Jamaica FSSA.xlsx', index = False, header = True)

Canada = Document ('Desktop/2019 Canada FSSA.docx')



for i in range(len(Canada.tables)):
    table = Canada.tables[i]
    data4= []
    keys = None
    row_data4 = None
    for j, row in enumerate(table.rows):
        text = (cell.text for cell in row.cells)
        if j == 0:
            keys = tuple(text)
            continue
        row_data4= dict(zip(keys,text))
        data4.append(row_data4)
    df4 = pd.DataFrame(data4)

df4.to_excel(r'Desktop/tables/Canada FSSA.xlsx', index = False, header = True)

USA = Document ('Desktop/2020 USA FSSA.docx')



for i in range(len(USA.tables)):
    table = USA.tables[i]
    data= []
    keys = None
    row_data = None
    for j, row in enumerate(table.rows):
        text = (cell.text for cell in row.cells)
        if j == 0:
            keys = tuple(text)
            continue
        row_data= dict(zip(keys,text))
        data.append(row_data)
    df5 = pd.DataFrame(data)



df5.to_excel(r'Desktop/tables/USA FSSA.xlsx', index = False, header = True)

