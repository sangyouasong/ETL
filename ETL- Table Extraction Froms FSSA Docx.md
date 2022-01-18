```python
#install python-docx library for extracting data from docx documents
!pip install python-docx
```


```python
#import libraries 
from docx.api import Document
from docx import Document 
import pandas as pd
```


```python
#find the path of targeted documents for table extraction 
!ls Desktop/*.docx
```

    [31mDesktop/2017 India FSSA.docx[m[m       Desktop/IMF interview.docx
    [31mDesktop/2017 New Zealand FSSA.docx[m[m Desktop/Macroeconimcs .docx
    [31mDesktop/2018 Jamaica FSSA.docx[m[m     Desktop/~$F interview.docx
    [31mDesktop/2019 Canada FSSA.docx[m[m      Desktop/~$croeconimcs .docx
    [31mDesktop/2020 USA FSSA.docx[m[m



```python
India = Document ('Desktop/2017 India FSSA.docx')
```


```python
India.tables
```




    [<docx.table.Table at 0x123be34d0>,
     <docx.table.Table at 0x123be35d0>,
     <docx.table.Table at 0x123be3650>,
     <docx.table.Table at 0x123be3750>,
     <docx.table.Table at 0x123be3890>,
     <docx.table.Table at 0x123be39d0>,
     <docx.table.Table at 0x123be3910>,
     <docx.table.Table at 0x123be3a50>,
     <docx.table.Table at 0x123be3b50>,
     <docx.table.Table at 0x123be3a10>,
     <docx.table.Table at 0x123be3d50>,
     <docx.table.Table at 0x123be3e90>,
     <docx.table.Table at 0x123be3e10>,
     <docx.table.Table at 0x123be3f50>,
     <docx.table.Table at 0x123be3cd0>,
     <docx.table.Table at 0x123bdbcd0>,
     <docx.table.Table at 0x123bdbed0>,
     <docx.table.Table at 0x123bdb7d0>,
     <docx.table.Table at 0x123bdbd90>,
     <docx.table.Table at 0x123bdb690>,
     <docx.table.Table at 0x123bdb450>,
     <docx.table.Table at 0x123bdb250>,
     <docx.table.Table at 0x123bdbfd0>,
     <docx.table.Table at 0x123bdb610>,
     <docx.table.Table at 0x123bdb890>,
     <docx.table.Table at 0x123bdbad0>,
     <docx.table.Table at 0x123bdb050>,
     <docx.table.Table at 0x123bdbe90>,
     <docx.table.Table at 0x123bdbc50>,
     <docx.table.Table at 0x123bdb1d0>,
     <docx.table.Table at 0x123bdbbd0>,
     <docx.table.Table at 0x123517510>,
     <docx.table.Table at 0x123517490>,
     <docx.table.Table at 0x123517950>,
     <docx.table.Table at 0x123517150>]




```python
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


```


```python
print(df1)
```

                                          Recommendations  \
    0                        Addressing system-wide risks   
    1   Enhance RBI monitoring of corporate indebtedne...   
    2   Improve the performance and financial strength...   
    3                          Financial sector oversight   
    4   Strengthen oversight of overseas operations of...   
    5   Enhance formal statutory basis for the autonom...   
    6   Tighten the definition of large and related-pa...   
    7   Enhance specialized expertise available to the...   
    8   Continue to strengthen coordination and inform...   
    9   Provide a lead supervisor with legal backing f...   
    10  Expedite passage of Insurance Law (Amendment) ...   
    11  Implement a corrective action ladder for insur...   
    12  Enact legislation to formalize the New Pension...   
    13  Systemic liquidity, crisis management, and saf...   
    14  Announce a timetable for the gradual reduction...   
    15  Strengthen resolution tools by granting strong...   
    16  Develop and periodically test arrangements to ...   
    
                                          Priority\n(H/M)  \
    0                        Addressing system-wide risks   
    1                                                   H   
    2                                                   H   
    3                          Financial sector oversight   
    4                                                   H   
    5                                                   M   
    6                                                   H   
    7                                                   H   
    8                                                   H   
    9                                                   H   
    10                                                  H   
    11                                                  H   
    12                                                  H   
    13  Systemic liquidity, crisis management, and saf...   
    14                                                  M   
    15                                                  H   
    16                                                  H   
    
                                               Time frame  \
    0                        Addressing system-wide risks   
    1                                                   S   
    2                                                   M   
    3                          Financial sector oversight   
    4                                                   M   
    5                                                   M   
    6                                                   M   
    7                                                   M   
    8                                                   S   
    9                                                   S   
    10                                                  S   
    11                                                  S   
    12                                                  S   
    13  Systemic liquidity, crisis management, and saf...   
    14                                                  M   
    15                                                  M   
    16                                                  M   
    
                                                   Status  
    0                        Addressing system-wide risks  
    1                                                   I  
    2                                                  PI  
    3                          Financial sector oversight  
    4                                                   I  
    5                                                   I  
    6                                                  PI  
    7                                                   I  
    8                                                   I  
    9                                                   I  
    10                                                  I  
    11                                                  I  
    12                                                  I  
    13  Systemic liquidity, crisis management, and saf...  
    14                                                 PI  
    15                                                 NI  
    16                                                 NI  



```python
df1.to_excel(r'Desktop/tables/India FSSA.xlsx', index = False, header = True)
```


```python
##Extracting table from 2017 New Zealand 
```


```python
NZ = Document ('Desktop/2017 New Zealand FSSA.docx')
```


```python
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


```


```python
df2.to_excel(r'Desktop/tables/NZ FSSA.xlsx', index = False, header = True)
```


```python
Jamaica = Document ('Desktop/2018 Jamaica FSSA.docx')

```


```python
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


```


```python
df3.to_excel(r'Desktop/tables/Jamaica FSSA.xlsx', index = False, header = True)
```


```python
Canada = Document ('Desktop/2019 Canada FSSA.docx')


```


```python
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
```


```python
df4.to_excel(r'Desktop/tables/Canada FSSA.xlsx', index = False, header = True)
```


```python
USA = Document ('Desktop/2020 USA FSSA.docx')


```


```python
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


```


```python
df5.to_excel(r'Desktop/tables/USA FSSA.xlsx', index = False, header = True)
```


```python

```
