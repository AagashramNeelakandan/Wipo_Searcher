##Code to Search whether a trademark available in Wipo site. It is takes the list of names in a text file and output the result to a Excel with Keywords Unique and Present

import xlwt 
from xlwt import Workbook
from openpyxl import load_workbook
import numpy as np
import os
from os import path

from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys


def load_names(start,end):
    file = open("Names_toCheck.txt","r")
    names = file.readlines()
    run_names = [i.strip('\n') for i in names[start:end]]
    return run_names


def check_names(start,end,num):
    names = load_names(start,end)
    filename = "Name_Details.xls"
    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')

    #Heading for excel
    sheet1.write(0, 0, 'Names')
    sheet1.write(0, 1, 'In Wipo')
   
    row=0
    col=0

    opts = Options()
    opts.headless = True # Operating in headless mode 
    browser = Firefox(options=opts)
    key1 = "Unique"
    key2 = "Present"
    actres=""
    
    for name in names:
        col = 0
        row = row+1
        sheet1.write(row,col, name)
        url =  'https://www3.wipo.int/branddb/en/?q={"searches":[{"te":"'+name+'","fi":"BRAND"}]}'
        browser.get(url)
        results = browser.find_elements_by_class_name('pagerPos')
        col = col+1
        if(len(results)==0):
            actres = key1
        else:
            actres = key2

        sheet1.write(row,col, actres)
        
        print(name+" -- "+actres)
        
        num = num-1
        if num<=0:
            browser.close()
            break
    wb.save(filename)       
           

def run_search(start,end):

    nos =end-start
    check_names(start,end,nos)
    
