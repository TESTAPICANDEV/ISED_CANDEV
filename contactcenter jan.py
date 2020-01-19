import pandas as pd
import numpy as np
import xlwt
from xlwt import Workbook
from csv import writer
from csv import reader
from collections import defaultdict


df = pd.read_excel(r"C:\Users\maanu\OneDrive\Desktop\Manav\candev\Contact Centre Data-20200119T062115Z-001\Contact Centre Data\Can Dev - Hackathon - Contact Centre Data - Calls July 2019.xlsx", sheet_name="Sheet1")
print (df)
# insert the name of the column as a string in brackets
topic = list(df['Topic']) 
resolution = list(df['Resolution'])
rfe = list(df['Reason for Enquiry'])
print(len(resolution))
print(len(rfe))
print(len(topic))
#resolution key
li = []
for i in range(len(resolution)):
    li.append((resolution[i],(topic[i], rfe[i])))
li.sort(key=lambda r:r[0])
completed=0
for i in range(len(li)):
    if li[i][0]=="Completed":
        completed+=1
        
partialres=0
for i in range(len(li)):
    if li[i][0]=="Partial Resolution + External" or li[i][0]=="Partial Resolution + Internal":
        partialres+=1
incomplete=0
for i in range(len(li)):
    if li[i][0]!="Completed" and li[i][0]!="Partial Resolution + External" and li[i][0]!="Partial Resolution + Internal":
        incomplete+=1
print("COmpleted",completed)
print("Partial resolution",partialres)
print("incomplete",incomplete)
        
'''
print(len(resolution))
analysis = {}
for i in range(len(resolution)):
    analysis[resolution[i]] = (topic[i], rfe[i])
    
print(analysis)
'''

wb=Workbook()
sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(0, 0, 'Topic')
sheet1.write(0, 1, 'Resolution')
sheet1.write(0, 2, 'Reason for Enquiry')
wb.save('xlwt contact data center resolution analysis july.xls')

r=1
for j in range(len(li)):
    sheet1.write(r, 0, li[j][1][0])
    r+=1
r=1

for j in range(len(li)):
    sheet1.write(r, 1, li[j][0])
    r+=1
r=1

for j in range(len(li)):
    sheet1.write(r, 2, li[j][1][1])
    r+=1
r=1
wb.save('xlwt contact data center resolution analysis july.xls')

cmplst=[]
for i in range(len(li)):
    if li[i][0]=="Completed":
        cmplst.append(li[i][1][0])
cmpctr=0
cmplst.sort()

ans=max(cmplst,key=cmplst.count)

            
incmplst=[]
for i in range(len(li)):
    if li[i][0]!="Completed" and li[i][0]!="Partial Resolution + External" and li[i][0]!="Partial Resolution + Internal":
        incmplst.append(li[i][1][0])
cmpctr=0
incmplst.sort()

ans2=max(incmplst,key=cmplst.count)
print("Most common unresolved topic",ans2)
cnt=0
unr=[]
for m in incmplst:
    if m==ans2:
        cnt+=1
print(ans2,"was unresolved",cnt,"number of times")


parcmp=[]
for i in range(len(li)):
    if li[i][0]=="Partial Resolution + External" or li[i][0]=="Partial Resolution + Internal":
        parcmp.append(li[i][1][0])
cmpctr=0
parcmp.sort()

ans3=ans2=max(parcmp,key=parcmp.count)

