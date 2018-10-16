
# coding: utf-8

# In[56]:


import pandas as pd
import re

DataSet=pd.read_excel('Testing 1.xlsx')
#WinItemName
#Mainframe Description
MainFrDes=DataSet['Mainframe Description']
my_dict = {}
Result = []
for Item in MainFrDes:
    output = re.findall('\s\d\S*\s[X]\s\d\S*\s[X]\s\d\S*\d|\d\S*\s[X]\s\S*\d|\s\d\S*[X]\S*[X]\S*\d\s|\d\S*[X]\S*\d|\d[-]\d[/]\d+|\d+[/]\d+|\d+[-][I][N]|\d+[I][N]|\d+[\'\"]|\s\d\s|^[1-9]\s',Item)
    #output= output.replace('i','')
    if output == []:
        Result.append('.')
    elif len(output)>0:
        Result.extend([' x '.join(output)])
  
Result1=[]
for x in Result:
    x=x.replace('"','')
    x=x.replace('\'','')
    x=x.replace('IN','')
    x=x.replace('X',' x ')                    
    Result1.append(x)

my_dict['key'] = Result1

ItemSize=pd.DataFrame(data = my_dict)

#print(ItemSize)

DataSet['Item Size MFD']=ItemSize
#print(DataSet)

excel_output = pd.ExcelWriter('Output 1.xlsx')
DataSet.to_excel(excel_output,'Sheet1',index=False)
excel_output.save()


# In[57]:


DataSet=pd.read_excel('Output 1.xlsx')
#WinItemName
#Mainframe Description
MainFrDes=DataSet['WinItemName']
my_dict = {}
Result = []
for Item in MainFrDes:
    output = re.findall('\d\S*\s[x]\s\d\S*\s[x]\s\d\S*\s[inf]|\d\S*\s[x]\s\d\S*\s[inf]|\d\S*\s[inf]',Item)
    if output == []:
        Result.append('.')
    elif len(output)>0:
        Result.extend([' x '.join(output)])
Result1=[]
for x in Result:
    x=x.replace('i','')
    x=x.replace('f','')
    Result1.append(x)
my_dict['key'] = Result1

ItemSize=pd.DataFrame(data = my_dict)

DataSet['Item Size WIN']=ItemSize


excel_output = pd.ExcelWriter('Output 2.xlsx')
DataSet.to_excel(excel_output,'Sheet1',index=False)
excel_output.save()

