# -*- coding: utf-8 -*-

#Created on Sun Jul 22 09:18:30 2018

#@author: akaliontzakis

from operator import itemgetter
import openpyxl, smtplib, sys
from datetime import datetime
from datetime import date
from email.mime.text import MIMEText
import pprint
from random import seed
from random import randint


def maketitles():

    #first roll to determine overall rarity. second roll to determine rarity of the beginning of title, if applicable. Third roll to determine rarity of the end of title, if applicable
    roll_1 = randint(0, 100)
    roll_2 = randint(0, 100)
    roll_3 = randint(0, 100)
    print(roll_1)
    print(roll_2)
    print(roll_3)
    titlelist = []
    finaltitles = []
    r1rarity = ''
    r2rarity = ''
    r3rarity = ''

    #list of excel workbooks to iterate through
    workbook = './titles.xlsx'
    #create list of tuples


    #iterate through workbooks

    wb = openpyxl.load_workbook(workbook)
        #select sheet
    sheet = wb['Sheet1']

    #iterate thru sheet
    for r in range(2, sheet.max_row + 1):
        title = sheet.cell(row=r, column=1).value
        location = sheet.cell(row=r, column=2).value
        rarity = sheet.cell(row=r, column=3).value
        titlelist.append((title,location,rarity))
        

    length = len(titlelist) - 1

    #low rarity title
    if roll_1 < 61:
        r1rarity = 'low'
        lowraritytitle = titlelist[randint(1,length)]
        print ('first try')
        print (lowraritytitle)
        while lowraritytitle[2] != "Low":
            print (lowraritytitle[2])
            while lowraritytitle[1] != "Middle":
            print(lowraritytitle[1])
            print('no good, trying again')
            lowraritytitle = titlelist[randint(1,length)]
            print (lowraritytitle)
        finaltitles.append((r1rarity, lowraritytitle[0]))
        #print (lowraritytitle[0])
        
    #medium rarity title   
    if roll_1 in range(61,95):
        r1rarity = 'medium'
        medraritytitle = ""
        medraritytitlebeg = titlelist[randint(1,length)]
        medraritytitlemid = titlelist[randint(1,length)]
        #determine the beginning of the title
        if roll_2 < 61:
            r2rarity = 'low'
            while medraritytitlebeg[2] != "Low" and medraritytitlebeg[1] != "Beginning":
                medraritytitlebeg = titlelist[randint(1,length)]
            medraritytitle = medraritytitlebeg[0]
        if roll_2 in range (61,95):
            r2rarity = 'medium'
            while medraritytitlebeg[2] != "Medium" and medraritytitlebeg[1] != "Beginning":
                medraritytitlebeg = titlelist[randint(1,length)]
            medraritytitle = medraritytitlebeg[0]
        if roll_2 in range (96,100):
            r2rarity = 'high'
            while medraritytitlebeg[2] != "Medium" and medraritytitlebeg[1] != "Beginning":
                medraritytitlebeg = titlelist[randint(1,length)]
            medraritytitle = medraritytitlebeg[0]     
            
            
        #determine the end of the title        
        while medraritytitlemid[2] != "Medium"  and medraritytitlemid[1] != "Middle":
            medraritytitlemid = titlelist[randint(1,length)]
        medraritytitle = medraritytitle + " " + medraritytitlemid[0]
        #print (medraritytitle)
        finaltitles.append((r1rarity, r2rarity, medraritytitle))    

    #high rarity title
    if roll_1 in range (96,100):
        r1rarity = 'high'    
        highraritytitle = ""
        highraritytitlebeg = titlelist[randint(1,length)]
        highraritytitlemid = titlelist[randint(1,length)]
        highraritytitleend = titlelist[randint(1,length)]
            #determine the beginning of the title
        if roll_2 < 61:
            r2rarity = 'low'
            while highraritytitlebeg[2] != "Low" and highraritytitlebeg[1] != "Beginning":
                highraritytitlebeg = titlelist[randint(1,length)]
            highraritytitle = highraritytitlebeg[0]
        if roll_2 in range (61,95):
            r2rarity = 'medium'
            while highraritytitlebeg[2] != "Medium" and highraritytitlebeg[1] != "Beginning":
                highraritytitlebeg = titlelist[randint(1,length)]
            highraritytitle = highraritytitlebeg[0]
        if roll_2 in range (96,100):
            r2rarity = 'high'
            while highraritytitlebeg[2] != "High" and highraritytitlebeg[1] != "Beginning":
                highraritytitlebeg = titlelist[randint(1,length)]
            highraritytitle = highraritytitlebeg[0]     
      
        #determine middle of title
        while highraritytitlemid[2] != "High" and highraritytitlemid[1] != "Middle":
            highraritytitlemid = titlelist[randint(1,length)]
        highraritytitle = highraritytitle + " " + highraritytitlemid[0]      
          
        #determine the end of the title      
        if roll_3 < 61:
            r3rarity = 'low'  
            while highraritytitleend[2] != "Low" and highraritytitleend[1] != "End":
                highraritytitleend = titlelist[randint(1,length)]
            highraritytitle = highraritytitle + " " + highraritytitleend[0] 
        if roll_3 in range (61,95):
            r3rarity = 'medium'  
            while highraritytitleend[2] != "Medium"  and highraritytitleend[1] != "End":
                highraritytitleend = titlelist[randint(1,length)]
            highraritytitle = highraritytitle + " " + highraritytitleend[0] 
        if roll_3 in range (96,100):
            r3rarity = 'high'  
            while highraritytitleend[2] != "High"  and highraritytitleend[1] != "End":
                highraritytitleend = titlelist[randint(1,length)]
            highraritytitle = highraritytitle + " " + highraritytitleend[0]    
        #print (highraritytitle)
        finaltitles.append((r1rarity, r2rarity, r3rarity, highraritytitle))  


    print(finaltitles)
#    for row in finaltitles:
 #       print ('%s | %s | %s | %s \n' % row)

    
for i in range(0,10):
    maketitles()