# -*- coding: utf-8 -*-
"""
Created on Sun Oct 20 14:57:59 2019

@author: 86171
"""

import PySimpleGUI as sg
import re
import datetime
import pyap
import pandas as pd
import difflib
import pgeocode

sg.SetOptions(
                 button_color = sg.COLOR_SYSTEM_DEFAULT
               , text_color = sg.COLOR_SYSTEM_DEFAULT
             )
sg.ChangeLookAndFeel('GreenTan')      

layout=[[sg.Text('Type Adress', size=(15, 1),auto_size_text=False, justification='left')
        ,sg.Input(key='_IN_',size=(20,1)),sg.Button('Find Supplier', button_color=('white', 'green'),size=(12, 1)),
        sg.Multiline(default_text='find result',size=(20,10),key='_Find Supplier_')],

        [sg.Text('Add New Adress', size=(15, 1),auto_size_text=False, justification='left'),
         sg.Input(key='_AddAdress_',size=(20,1))],
        
        [ sg.Text('Add New Supplier', size=(15, 1),auto_size_text=False, justification='left'),
         sg.Input(key='_AddSupplier_',size=(20,1))],    
         
        [sg.Button('Add Data', button_color=('white', 'green'),size=(14, 1)),
        sg.Multiline(default_text='Add Result',size=(18,5),key='_Add Result_')],
        
        [sg.Cancel()] ]
window = sg.Window('Fraser H').Layout(layout) 


def find_code(postcode1,code_list):
    dist = pgeocode.GeoDistance('AU')
    #postcode1 = '4208'
    #postcode2 = ['4209','3350','3355','3550','3175']
    distance_list = []
    for tempcode in code_list:
        distance=dist.query_postal_code(postcode1, tempcode)
        print(distance)
        distance_list.append(distance)
    m = min(distance_list)
    print(distance_list)

    dis = []
    for i in range(len(distance_list)):
        if distance_list[i] != m:
            dis.append(distance_list[i])
            
    print("Distance_first",distance_list)
    print("second : ",dis)
    m_second = min(dis)
    print("m_sec", m_second)
    
    i = 0
    n = 0
    for temp in code_list:
        if dist.query_postal_code(postcode1,temp) == m :
            n = i
            resultcode = temp
            print(i,temp)
            distance = round(distance_list[n],2)
        i += 1
    i = 0
    n_second = 0
    for temp in code_list:
        if dist.query_postal_code(postcode1,temp) == m_second:
            n_second = i
            result_code_second = temp
            distance_second = round(distance_list[n_second],2)
            print(i,temp)
        i += 1
    print(n,resultcode,distance,n_second,result_code_second,distance_second)
    return n,resultcode,distance,n_second,result_code_second,distance_second


def find_supplier(code):
    file_name = "source.xlsx"
    df = pd.read_excel(file_name)
    print(df)
    pipeline = ''
    i = 0
    code_list = []
    for i in range(df.shape[0]): 
        words=df['address'][i].split(' ')
        words.append(df['address'][i])
        #result=difflib.get_close_matches(code,words)
        tempcode = df['address'][i][-4:]
        code_list.append(tempcode)
        #print('result : ',result)
        #if len(result) > 0:
        #    pipeline += str(i+1) + " : " +  df['supplier'][i] + '\n'
    try:
        n,resultcode,distance,n_s,resultcode_sec,distance_sec = find_code(code,code_list)
    #distance=dist.query_postal_code(code, resultcode)
        supplier = df['supplier'][n]
        supplier_sec = df['supplier'][n_s]
        pipeline +=   "\n\nSupplier : " + supplier  + '\n' + "Distance(" + code + "," + resultcode + ') = ' + str(distance) + 'Km\n\n' + "\nSecondSupplier : " + supplier_sec + "\n" + "Distance(" + code + "," + resultcode_sec+") = " + str(distance_sec) + 'Km'
    except:
        pipeline = "Please retype vaild post code or address"
    window.FindElement('_Find Supplier_').Update(pipeline)
    

def strToTxt(resultfileName,out_Text):
    with open(resultfileName + '.xlsx','a',encoding='utf-8') as f:
        f.write(out_Text)
        f.write('\n')
        
def adddata(rows,filename):
    from openpyxl import Workbook

    book = Workbook()
    sheet = book.active
    for row in rows:
        sheet.append(row)

    book.save('filename')
    
while True:                 # Event Loop  
    event, values = window.Read()
    print(event,values)
  
    if event is None or event == 'Exit':  
        break 
    
    if event == 'Find Supplier':
        print(values['_IN_'])
        post_code = values['_IN_'][-4:] 
        find_supplier(post_code)
        
    if event == 'Add Data':
        print(values['_AddAdress_'])
        print(values['_AddSupplier_'])
        if (values['_AddAdress_'] != '' and values['_AddSupplier_'] != ''):
            print('yes')
            df = pd.read_excel('source.xlsx')
            #print(list(df['address'].values).append(values['_AddAdress_']))
            n = df.shape[0]
            df['address'][n+1] = str(values['_AddAdress_'])
            df['supplier'][n+1] = str(values['_AddSupplier_'])
            df1 = pd.DataFrame(list(zip(df['supplier'], df['address'])), 
               columns =['supplier', 'address'])
            print(df['address'][n+1])
            df1.to_excel('source.xlsx')
            window.FindElement('_Add Result_').Update('Updated!')
        #find_supplier(post_code)
