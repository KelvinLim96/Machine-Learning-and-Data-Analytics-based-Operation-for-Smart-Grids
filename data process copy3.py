import pandas as pd
import numpy as np
import xlrd
import os 
import logging
import sys
import xlwt
from openpyxl import load_workbook
book = xlwt.Workbook() 
sheet5 = book.add_sheet('sheet5',cell_overwrite_ok=True)
data = pd.read_excel('C:\\Users\\65837\\Desktop\\original3.xlsx')

print(float(data.iloc[2,[2]]))
#print(float(data.iloc[2,[7]]))
i=0
time = 0
count=0
n=0
for row in range(2,8786,720):
    sheet5.write(1,2,'EMB_P')
    while time < 24:
        for column in range(7,73,13):
            if column == 7:
                while(float(data.iloc[row+time-1+n,[column]])<0 or float(data.iloc[row+time-1+n,[column]])>80):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 20:
                while(float(data.iloc[row+time-1+n,[column]])<100 or float(data.iloc[row+time-1+n,[column]])>300):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0           
            if column == 33:
                while(float(data.iloc[row+time-1+n,[column]])<80 or float(data.iloc[row+time-1+n,[column]])>500):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 46:
                while(float(data.iloc[row+time-1+n,[column]])<50 or float(data.iloc[row+time-1+n,[column]])>600):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 59:
                while(float(data.iloc[row+time-1+n,[column]])<150 or float(data.iloc[row+time-1+n,[column]])>300):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 72:
                while(float(data.iloc[row+time-1+n,[column]])<50 or float(data.iloc[row+time-1+n,[column]])>200):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
        sheet5.write((int(row/720))*24+time+2,2,i)
        time = time+1
        i = 0
    time = 0
    
for row in range(2,8786,720):
    sheet5.write(1,3,'EMB_Q')
    while time < 24:
        for column in range(6,72,13):
            if column == 6:
                while(float(data.iloc[row+time-1+n,[column]])<-20 or float(data.iloc[row+time-1+n,[column]])>10):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 19:
                while(float(data.iloc[row+time-1+n,[column]])<-20 or float(data.iloc[row+time-1+n,[column]])>20):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0           
            if column == 32:
                while(float(data.iloc[row+time-1+n,[column]])<-60 or float(data.iloc[row+time-1+n,[column]])>60):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 45:
                while(float(data.iloc[row+time-1+n,[column]])<-50 or float(data.iloc[row+time-1+n,[column]])>50):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 58:
                while(float(data.iloc[row+time-1+n,[column]])<-50 or float(data.iloc[row+time-1+n,[column]])>50):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 71:
                while(float(data.iloc[row+time-1+n,[column]])<0 or float(data.iloc[row+time-1+n,[column]])>80):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
        sheet5.write((int(row/720))*24+time+2,3,i)
        time = time+1
        i = 0
    time = 0

for row in range(2,8786,720):
    sheet5.write(1,4,'NYA_P')
    while time < 24:
        for column in range(85,112,13):
            if column == 85:
                while(float(data.iloc[row+time-1+n,[column]])<1 or float(data.iloc[row+time-1+n,[column]])>30):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 98:
                while(float(data.iloc[row+time-1+n,[column]])<-80 or float(data.iloc[row+time-1+n,[column]])>80):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 111:
                while(float(data.iloc[row+time-1+n,[column]])<-600 or float(data.iloc[row+time-1+n,[column]])>150):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0 
        sheet5.write((int(row/720))*24+time+2,4,i)
        time = time+1
        i = 0
    time = 0

for row in range(2,8786,720):
    sheet5.write(1,5,'NYA_Q')
    while time < 24:
        for column in range(84,111,13):
            if column == 84:
                while(float(data.iloc[row+time-1+n,[column]])<1 or float(data.iloc[row+time-1+n,[column]])>5):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 97:
                while(float(data.iloc[row+time-1+n,[column]])<-20 or float(data.iloc[row+time-1+n,[column]])>50):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 110:
                while(float(data.iloc[row+time-1+n,[column]])<-20 or float(data.iloc[row+time-1+n,[column]])>50):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0    
        sheet5.write((int(row/720))*24+time+2,5,i)
        time = time+1
        i = 0
    time = 0

book.save('sheet5.xls')
            
    
        






    
'''sheet2.write(0,1,str('NEC EMS1_IN1'))
    sheet2.write(0,2,str('NEC EMS1_IN2'))
    sheet2.write(0,3,str('NEC EMS1_IN3'))
    sheet2.write(0,4,str('NEC MSB1_IN1'))
    sheet2.write(0,5,str('NEC MSB1-IN2'))
    sheet2.write(0,6,str('NEC MSB1-IN3'))
    sheet2.write(0,7,str('CANT2 MSB1_IN1'))
    sheet2.write(0,8,str('CANT2 MSB1_IN2'))
    sheet2.write(0,9,str('SPMS EMSB1_IN1'))
    sheet2.write(0,10,str('SPMS MSB1_IN1'))
    sheet2.write(0,11,str('SPMS MSB1_IN2'))
    sheet2.write(0,12,str('SPMS MSB1_IN3'))
    sheet2.write(0,11,str('SPMS MSB2_IN1'))
    sheet2.write(0,12,str('SPMS MSB2_IN2'))
    
    sheet2.write(0,13,str('SPMS MSB1_IN2'))
    sheet2.write(0,14,str('SPMS MSB1_IN3'))
    sheet2.write(0,15,str('SPMS MSB1_IN2'))
    sheet2.write(0,16,str('SPMS MSB1_IN3'))'''
   
