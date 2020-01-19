\import pandas as pd
import numpy as np
from csv import writer
from csv import reader
import xlwt
from xlwt import Workbook
wb = Workbook()
# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')
df=pd.read_csv("CanadaBusiness  tweet_activity_metrics - Apr 2019 -.csv")
print(df.head())
hasht=""
li=[]
l1=[]
li2=[]
i=0
for row in df["Tweet text"]:
    li=row.split()
    for i in range(len(li)):
        if "#" in li[i]:
            sample=li[i]
            length=len(sample)
            if sample[length-1] in [":",".",",",""]:
                sample[0:-1]
            li2.append(sample)
    print(row)
print(li2)

sheet1.write(0, 1, 'Entrepreneurship')
sheet1.write(0, 2, 'Clean Technology')
sheet1.write(0, 3, 'Business/Tips')
sheet1.write(0, 4, 'DYK')
sheet1.write(0, 5, 'Taxes')
sheet1.write(0, 6, 'Budget')
sheet1.write(0, 7, 'Coding and Developement')
sheet1.write(0, 8, 'Innovation Canada')
sheet1.write(0, 9, 'Invest in Canada')
sheet1.write(0, 10, 'Export/Import')
r=1
c=1
entre=[]
ct=[]
bt=[]
dyk=[]
tax=[]
b=[]
cd=[]
ic=[]
inc=[]
tra=[]

wb.save('xlwt Classified Tweets April.xls')
for row in df['Tweet text']:
    l1=row.split()
    print(l1)
    for i in range(len(l1)):
        if l1[i] in ["#youngentrepreneurs","#youngentrepreneurs","#startup","#entrepreneurs","#Womenentrepreneurs","#startups!"]:
            entre.append(row)
            print(entre)
        if l1[i] in ["#cleantech","#environment"]:
            ct.append(row)
        if l1[i] in ["#business","#SMEs","#BizTip","#business?","#Cdnbiz","#smallbiz","#cooperatives","#biztips"]:
            bt.append(row)
        if l1[i] in ["#DYK"]:
            dyk.append(row)
        if l1[i] in ["#taxschemes","#CdnTax","#taxes","#taxtip","taxes?"]:
            tax.append(row)
        if l1[i] in ["#YourBudget2019"]:
            b.append(row)
        if l1[i] in ["#CanCode","#digitalskills","#coders","developers","GCAPIstore","#GCdigital"]:
            cd.append(row)
        if l1[i] in ["#InnovationCanada","#newtech","#innovation","#innovations"]:
            ic.append(row)
        if l1[i] in ["#InvestInCanada"]:
            inc.append(row)
        if l1[i] in ["#trades","#export","#trade"]:
            tra.append(row)
print (entre)
for j in entre:
    sheet1.write(r, 1, j)
    r+=1
r=1

for j in ct:
    sheet1.write(r, 2, j)
    r+=1
r=1
for j in bt:
    sheet1.write(r, 3, j)
    r+=1
r=1
for j in dyk:
    sheet1.write(r, 4, j)
    r+=1
r=1
for j in tax:
    sheet1.write(r, 5, j)
    r+=1
r=1
for j in b:
    sheet1.write(r, 6, j)
    r+=1
r=1
for j in cd:
    sheet1.write(r, 7, j)
    r+=1
r=1
for j in ic:
    sheet1.write(r, 8, j)
    r+=1
r=1
for j in inc:
    sheet1.write(r, 9, j)
    r+=1
r=1
for j in tra:
    sheet1.write(r, 10, j)
    r+=1
r=1
wb.save('xlwt Classified Tweets April.xls')
