# -*- coding: utf-8 -*-
"""
Created on Fri Feb 19 15:54:45 2021

@author: playe
"""

import os
import calendar
import datetime
from copy import copy
from datetime import date, timedelta
from openpyxl.formula.translate import Translator
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment

declared_roi = 1.5
issued = "AP Letter Issued"
cancelled = "AP Letter Cancelled"
purchased = "Home Purchased"

#######################
### Get directories ###
#######################
dir = os.getcwd()
dir_pre = os.path.join(dir, "Before")
dir_post = os.path.join(dir, "After")
dir_error = os.path.join(dir, "error_log.xlsx")

if os.path.isfile(dir_error):
    print ("File exist")
    os.remove(dir_error)
else:
    print ("File not exist")
error_wb = Workbook()
error_ws = error_wb.active
principal_wb = Workbook()
principal_ws = principal_wb.active
dir_list = sorted(os.listdir(dir_pre))
print(dir)
##################################
### Get current quarter months ###
##################################

input_day = date(year = 2020, month = 10, day = 5)
today = input_day.day
month = input_day.month
year = input_day.year
m = 10
quarter = ((m-1)//3 + 1)
pstr = str(year) + "Q" + str(quarter) + "principal_column.xlsx"
dir_principal = os.path.join(dir, pstr)
roi_string = "ROI " + str(input_day.year) + "Q" + str(quarter) + ": " + "{:.2f}%".format(declared_roi)
print(roi_string)
if (quarter == 1):
    month1 = 1
    month2 = 2
    month3 = 3
elif (quarter == 2):
    month1 = 4
    month2 = 5
    month3 = 6
elif (quarter == 3):
    month1 = 7
    month2 = 8
    month3 = 9
elif (quarter == 4):
    month1 = 10
    month2 = 11
    month3 = 12
else :
    print("Invalid Quarter")
    
start_of_m1 = date(year = input_day.year, month = month1, day = 1)
start_of_m2 = date(year = input_day.year, month = month2, day = 1)
start_of_m3 = date(year = input_day.year, month = month3, day = 1)
end_of_m1 =  date(year = input_day.year, month = month1, day = calendar.monthrange(input_day.year, month1)[1])
end_of_m2 = date(year = input_day.year, month = month2, day = calendar.monthrange(input_day.year, month2)[1])
end_of_m3 = date(year = input_day.year, month = month3, day = calendar.monthrange(input_day.year, month3)[1])
end_of_quarter = end_of_m3


print("Month 1:", calendar.month_name[month1], "Last day:", calendar.monthrange(input_day.year, month1)[1])
print("Month 2:", calendar.month_name[month2], "Last day:", calendar.monthrange(input_day.year, month2)[1])
print("Month 3:", calendar.month_name[month3], "Last day:", calendar.monthrange(input_day.year, month3)[1])    
print("End of Quarter:", end_of_quarter)
print(end_of_quarter)

#######################
### File processing ###
#######################
COUNT = 1
def increment():
    global COUNT
    COUNT = COUNT+1
p_COUNT = 1
def p_increment():
    global p_COUNT
    p_COUNT = p_COUNT+1
    
def partial_roi(ws, file, start_of_m, end_of_m, i):
    in_range = 0
    cell = ws.cell(row = i, column = 1).value
    minimum = ws.cell(row = i, column = 7).value
    current_date = date(year = cell.year, month = cell.month, day = cell.day)
    j = i + 1
    while(current_date <= end_of_m):
        try:
            if(in_range == 0):
                minimum = ws.cell(row = i, column = 7).value
            cell = ws.cell(row = i, column = 1).value
            nextcell = ws.cell(row = j, column = 1).value
            current_date = date(year = cell.year, month = cell.month, day = cell.day)
            next_date = date(year = nextcell.year, month = nextcell.month, day = nextcell.day)
            ### comparison error case
            if(next_date < current_date):
                print("Unordered dates:", current_date, "and", next_date)
                error_ws.cell(row = COUNT, column = 1).value = file 
                error_ws.cell(row = COUNT, column = 2).value = "Dates not in order"
                increment()
                error_wb.save(dir_error)
                i += 1
                j += 1
                continue
                # return {'partial_roi' : 0, 'error' : 1, 'r' : i, 'EoF': 0}
                # return 0, 1, i, 0
            ### must continue until we reach m
            if(next_date <= start_of_m):
                i += 1
                j += 1
                continue
            elif((next_date > start_of_m) & (next_date <= end_of_m)):
                ###continue calculation...
                print("In range")
                in_range = 1
                if(minimum >= ws.cell(row = j, column = 7).value):
                    minimum = ws.cell(row = j, column = 7).value
                    i = j
                j+=1
                print(minimum)
                continue
            ### outside of range, take last minimum
            elif(next_date > end_of_m):
                # return{'partial_roi' : minimum, 'error' : 0, 'r' : i, 'EoF' : 0}
                return minimum, 0, i, 0
        except AttributeError:
            if((ws.cell(row = i, column = 7).value != None) & (ws.cell(row = j, column = 7).value != None)):
            ### missing dates should be somewhat frequent in the beginning
                if(in_range == 1):
                    print("Critical error: Missing date in quarter")
                    error_ws.cell(row = COUNT, column = 1).value = file 
                    error_ws.cell(row = COUNT, column = 2).value = "Missing date in quarter"
                    increment()
                    error_wb.save(dir_error)
                    # return{'partial_roi' : minimum, 'error' : 1, 'r' : i, 'EoF' : 1}
                    return minimum, 1, i, 1
                else:
                    print("Missing date")
                    i += 1
                    j += 1
                    continue
            else:
                print("End of file, returning roi")
                # return{'partial_roi' : minimum, 'error' : 0, 'r' : i, 'EoF' : 1}
                minimum = ws.cell(row = j-1, column = 7).value
                return minimum, 0, i, 1
            
    
    print("Done partial roi")
    # return{'partial_roi' : minimum, 'error' : 0, 'r' : i, 'EoF' : 0}
    return minimum, 0, i, 0

def write_principal_column(file, p1, p2, p3, roi):
    principal_ws.cell(row = p_COUNT, column = 1).value = file
    principal_ws.cell(row = p_COUNT, column = 2).value = p1
    principal_ws.cell(row = p_COUNT, column = 3).value = p2
    principal_ws.cell(row = p_COUNT, column = 4).value = p3
    principal_ws.cell(row = p_COUNT, column = 5).value = roi
    principal_wb.save(dir_principal)
def process_file(infile, outfile, file):
    out_of_range = 0
    r = 9
    p1 = 0
    p2 = 0
    p3 = 0
    ap_issued = 0
    ap_cancelled = 0

    os.chdir(dir_pre)
    book = load_workbook(infile, data_only = True)
    ws = book.active
    
    cell = ws.cell(row = r, column = 1).value
    try:
        current_date = date(year = cell.year, month = cell.month, day = cell.day)
        if(current_date > end_of_m3):
            out_of_range = 1
            print("First date is after the current quarter. Continuing...")
            error_ws.cell(row = COUNT, column = 1).value = file 
            error_ws.cell(row = COUNT, column = 3).value = "After quarter"
            increment()
            error_wb.save(dir_error)
            return
    except :
        print("Missing date in first row")
        error_ws.cell(row = COUNT, column = 1).value = file 
        error_ws.cell(row = COUNT, column = 3).value = "Missing first row"
        increment()
        error_wb.save(dir_error)
        return
    
    
    p1, error, r, EoF = partial_roi(ws, file, start_of_m1, end_of_m1, r)
    if(error):
        print("error: check log file")
        return
    if(EoF):
        p2 = p1
        p3 = p1
    else:
        p2, error, r, EoF = partial_roi(ws, file, start_of_m2, end_of_m2, r)
        if(error):
            print("error: check log file")
            return
        if(EoF):
            p3 = p2
        else:
            p3, error, r, EoF = partial_roi(ws, file, start_of_m3, end_of_m3, r)
            if(error):
                print("error: check log file")
                return
    print("P1:",p1)
    print("P2:",p2)
    print("P3:",p3)
    roi = (p1 + p2 + p3) * declared_roi/(3*100)
    write_principal_column(file, p1, p2, p3, roi)
    p_increment()
    book.close()
    
    
for file in sorted( filter(lambda x: not (x.startswith('~') or x.startswith('.')), dir_list) ):
#file = "0733C2.xlsx"
    file_in = os.path.join(dir_pre, file)
    file_out = os.path.join(dir_post, file)
    #file_in = file_in.replace(os.sep, '/')
    print(file_in)
    process_file(file_in, file_out, file)   
            
    
        
        
        























