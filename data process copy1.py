import pandas as pd
import numpy as np
import xlrd
import os 
import logging
import sys
import xlwt
from openpyxl import load_workbook
book = xlwt.Workbook() 
sheet3 = book.add_sheet('sheet3',cell_overwrite_ok=True)
data = pd.read_excel('C:\\Users\\65837\\Desktop\\original1.xlsx')

print(float(data.iloc[2,[2]]))
#print(float(data.iloc[2,[7]]))
i=0
time = 0
count=0
n=0
for row in range(2,8786,720):
    sheet3.write(1,2,'N1_3_P')
    while time < 24:
        for column in range(6,33,13):
            if column == 6:
                while(float(data.iloc[row+time-1+n,[column]])<30 or float(data.iloc[row+time-1+n,[column]])>200):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 19:
                while(float(data.iloc[row+time-1+n,[column]])<150 or float(data.iloc[row+time-1+n,[column]])>700):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0           
            if column == 32:
                while(float(data.iloc[row+time-1+n,[column]])<100 or float(data.iloc[row+time-1+n,[column]])>650):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            
        sheet3.write((int(row/720))*24+time+2,2,i)
        time = time+1
        i = 0
    time = 0
    
for row in range(2,8786,720):
    sheet3.write(1,3,'N1_3_Q')
    while time < 24:
        for column in range(5,32,13):
            if column == 5:
                while(float(data.iloc[row+time-1+n,[column]])<-5 or float(data.iloc[row+time-1+n,[column]])>80):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 18:
                while(float(data.iloc[row+time-1+n,[column]])<-10 or float(data.iloc[row+time-1+n,[column]])>400):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0           
            if column == 31:
                while(float(data.iloc[row+time-1+n,[column]])<-10 or float(data.iloc[row+time-1+n,[column]])>300):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
        sheet3.write((int(row/720))*24+time+2,3,i)
        time = time+1
        i = 0
    time = 0

for row in range(2,8786,720):
    sheet3.write(1,4,'N_2_P')
    while time < 24:
        for column in range(45,72,13):
            if column == 45:
                while(float(data.iloc[row+time-1+n,[column]])<-80 or float(data.iloc[row+time-1+n,[column]])>200):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 58:
                while(float(data.iloc[row+time-1+n,[column]])<0 or float(data.iloc[row+time-1+n,[column]])>200):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 71:
                while(float(data.iloc[row+time-1+n,[column]])<20 or float(data.iloc[row+time-1+n,[column]])>250):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0  
            
        sheet3.write((int(row/720))*24+time+2,4,i)
        time = time+1
        i = 0
    time = 0

for row in range(2,8786,720):
    sheet3.write(1,5,'N_2_Q')
    while time < 24:
        for column in range(44,71,13):
            if column == 44:
                while(float(data.iloc[row+time-1+n,[column]])<-80 or float(data.iloc[row+time-1+n,[column]])>200):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 57:
                while(float(data.iloc[row+time-1+n,[column]])<-50 or float(data.iloc[row+time-1+n,[column]])>150):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 70:
                while(float(data.iloc[row+time-1+n,[column]])<-50 or float(data.iloc[row+time-1+n,[column]])>100):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0   
            
        sheet3.write((int(row/720))*24+time+2,5,i)
        time = time+1
        i = 0
    time = 0

for row in range(2,8786,720):
    sheet3.write(1,6,'N_2_1_P')
    while time < 24:
        for column in range(84,111,13):
            if column == 84:
                while(float(data.iloc[row+time-1+n,[column]])<-5 or float(data.iloc[row+time-1+n,[column]])>100):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 97:
                while(float(data.iloc[row+time-1+n,[column]])<-5 or float(data.iloc[row+time-1+n,[column]])>400):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 110:
                while(float(data.iloc[row+time-1+n,[column]])<50 or float(data.iloc[row+time-1+n,[column]])>800):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0 
            
        sheet3.write((int(row/720))*24+time+2,6,i)
        time = time+1
        i = 0
    time = 0

for row in range(2,8786,720):
    sheet3.write(1,7,'N2_1_Q')
    while time < 24:
        for column in range(83,110,13):
            if column == 83:
                while(float(data.iloc[row+time-1+n,[column]])<-10 or float(data.iloc[row+time-1+n,[column]])>20):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 96:
                while(float(data.iloc[row+time-1+n,[column]])<-30 or float(data.iloc[row+time-1+n,[column]])>100):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 109:
                while(float(data.iloc[row+time-1+n,[column]])<0 or float(data.iloc[row+time-1+n,[column]])>400):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0 
        sheet3.write((int(row/720))*24+time+2,7,i)
        time = time+1
        i = 0
    time = 0

for row in range(2,8786,720):
    sheet3.write(1,8,'SBS_P')
    while time < 24:
        for column in range(123,189,13):
            if column == 123:
                while(float(data.iloc[row+time-1+n,[column]])<50 or float(data.iloc[row+time-1+n,[column]])>500):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 136:
                while(float(data.iloc[row+time-1+n,[column]])<10 or float(data.iloc[row+time-1+n,[column]])>300):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 149:
                while(float(data.iloc[row+time-1+n,[column]])<100 or float(data.iloc[row+time-1+n,[column]])>600):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0 
            if column == 162:
                while(float(data.iloc[row+time-1+n,[column]])<200 or float(data.iloc[row+time-1+n,[column]])>800):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 175:
                while(float(data.iloc[row+time-1+n,[column]])<250 or float(data.iloc[row+time-1+n,[column]])>800):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 188:
                while(float(data.iloc[row+time-1+n,[column]])<500 or float(data.iloc[row+time-1+n,[column]])>1100):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0               
        sheet3.write((int(row/720))*24+time+2,8,i)
        time = time+1
        i = 0
    time = 0

for row in range(2,8786,720):
    sheet3.write(1,9,'SBS_Q')
    while time < 24:
        for column in range(122,188,13):
            if column == 122:
                while(float(data.iloc[row+time-1+n,[column]])<40 or float(data.iloc[row+time-1+n,[column]])>150):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 135:
                while(float(data.iloc[row+time-1+n,[column]])<10 or float(data.iloc[row+time-1+n,[column]])>80):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 148:
                while(float(data.iloc[row+time-1+n,[column]])<50 or float(data.iloc[row+time-1+n,[column]])>150):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0 
            if column == 161:
                while(float(data.iloc[row+time-1+n,[column]])<100 or float(data.iloc[row+time-1+n,[column]])>200):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 174:
                while(float(data.iloc[row+time-1+n,[column]])<30 or float(data.iloc[row+time-1+n,[column]])>150):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
            if column == 187:
                while(float(data.iloc[row+time-1+n,[column]])<100 or float(data.iloc[row+time-1+n,[column]])>300):
                    n=n+1
                i = i+float(data.iloc[row+time-1+n,[column]])
                n=0
        sheet3.write((int(row/720))*24+time+2,9,i)
        time = time+1
        i = 0
    time = 0
book.save('sheet3.xls')
            
    
        






    
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
   
