# -*- coding: utf-8 -*-
"""
Created on Fri Feb 19 15:54:45 2021

@author: Usama
"""

import os
import calendar
import datetime
import sys
import tkinter as tk
from copy import copy
from datetime import date, timedelta
from openpyxl.formula.translate import Translator
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment

declared_roi = 1.5
issued = "ap letter issued"
cancelled = "ap letter cancelled"
purchased = "home purchased"

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
today = date.today()
tempqtr = ((today.month-1)//3 + 1)
button_exit = 0
def quitButton():
    global button_exit
    button_exit = 1
    master.destroy()
    
def startButton():
    # today = datetime.date(datetime.now())
    global quarter
    global year
    global dir_pre
    global dir_post
    global dir_list
    global write_principal
    global write_statements
    write_principal, write_statements = getBool()
    q,y,inp,outp = getText()
    try:
        quarter = int(q)
        year = int(y)
        if((quarter <= 0)|(quarter > 4)):
            raise ValueError
        dir_pre = inp
        dir_post = outp
        if(year > today.year):
            print("Invalid year, please try again")
        elif((year == today.year) & (tempqtr <= quarter)):
            print("Invalid quarter, please try again")
        else:
            print("Continuing...")    
            if(os.path.exists(dir_pre)):
                print("continuing...")
                if not os.listdir(dir_pre):
                    print("directory is empty")
                else:
                    print("continuing...")
                    dir_list = sorted(os.listdir(dir_pre))
                    master.destroy()     
            else:
                print("no such path")
    except ValueError:
        print("Invalid quarter or year, please try again")

def getBool():
    return var1.get(), var2.get()
def getText():
    global e1
    global e2
    global e3
    global e4
    print(var1.get())
    q = e1.get()
    y = e2.get()
    inp= e3.get()
    outp= e4.get()
    return q,y,inp,outp

#Setting root

master = tk.Tk()
frame = tk.Frame(master)
frame.pack()

#Creating checkbox variables
var1 = tk.BooleanVar(value=True)
var2 = tk.BooleanVar(value=False)
if (tempqtr == 1):
    entryq = 4
    entryy = today.year-1
else:
    entryq = tempqtr
    entryy = today.year
#Creating and setting frames
middleFrame = tk.Frame(master)
middleFrame.pack(side = "bottom")
bottomFrame1 = tk.Frame(middleFrame)
bottomFrame1.pack(side = "bottom")
bottomFrame2 = tk.Frame(bottomFrame1)
bottomFrame2.pack(side = "bottom")
bottomFrame3 = tk.Frame(bottomFrame2)
bottomFrame3.pack(side = "bottom") 
bottomFrame4 = tk.Frame(bottomFrame3)
bottomFrame4.pack(side = "bottom") 

#Creating and setting top and middle Labels and Entries
tk.Label(frame, text = "ROI calculator").pack(side = "top")
tk.Label(middleFrame, text="Quarter").pack(side = "left")
e1 = tk.Entry(middleFrame, width = 1)
e1.insert(0, entryq)
e1.pack(side = "left")
tk.Label(middleFrame, text="Year").pack(side = "left")
e2 = tk.Entry(middleFrame, width = 4)
e2.insert(0, entryy)
e2.pack(side = "left")

#Creating and setting Input and Output paths
tk.Label(bottomFrame1, text = "Input Path").pack(side = "left")
e3 = tk.Entry(bottomFrame1, width = 50)
e3.insert(0, dir_pre)
e3.pack(side = "right")
tk.Label(bottomFrame2, text = "Output Path").pack(side = "left")
e4 = tk.Entry(bottomFrame2, width = 50)
e4.insert(0, dir_post)
e4.pack(side = "right")

#Creating and setting Checkboxes and Start/Exit buttons
c1 = tk.Checkbutton(bottomFrame3, text = "Write Principal Column", variable = var1).pack(side = "top")
c2 = tk.Checkbutton(bottomFrame3, text = "Write Statements", variable = var2).pack(side = "left")
tk.Button(bottomFrame4, text = "Exit", padx = 10, command = quitButton).pack(side = "left")
tk.Button(bottomFrame4, text = "Start", padx = 10, command = startButton).pack(side = "right")

#Quitting after mainloop
tk.mainloop()
master.quit()

if(button_exit == 1):
    print("Exiting")
    sys.exit()
pstr = str(year) + "Q" + str(quarter) + "principal_column.xlsx"
dir_principal = os.path.join(dir, pstr)
roi_string = "ROI " + str(year) + "Q" + str(quarter) + ": " + "{:.2f}%".format(declared_roi)
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
    sys.exit()
    
start_of_m1 = date(year = year, month = month1, day = 1)
start_of_m2 = date(year = year, month = month2, day = 1)
start_of_m3 = date(year = year, month = month3, day = 1)
end_of_m1 =  date(year = year, month = month1, day = calendar.monthrange(year, month1)[1])
end_of_m2 = date(year = year, month = month2, day = calendar.monthrange(year, month2)[1])
end_of_m3 = date(year = year, month = month3, day = calendar.monthrange(year, month3)[1])
end_of_quarter = end_of_m3


print("Month 1:", calendar.month_name[month1], "Last day:", calendar.monthrange(year, month1)[1])
print("Month 2:", calendar.month_name[month2], "Last day:", calendar.monthrange(year, month2)[1])
print("Month 3:", calendar.month_name[month3], "Last day:", calendar.monthrange(year, month3)[1])    
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
    
def partial_roi(ws, file, start_of_m, end_of_m, i, ap):
    ap_issued = ap
    in_range = 0
    missing_date_error = 0
    unordered_date_error = 0
    cell = ws.cell(row = i, column = 1).value
    minimum = ws.cell(row = i, column = 7).value
    current_date = date(year = cell.year, month = cell.month, day = cell.day)
    if(current_date > start_of_m):
        if(i == 9):
            minimum = 0
        return minimum, 0, i, 0, ap_issued
    j = i + 1
    while(current_date <= end_of_m):
        try:
            cell = ws.cell(row = i, column = 1).value
            next_cell = ws.cell(row = j, column = 1).value
            current_date = date(year = cell.year, month = cell.month, day = cell.day)
            next_date = date(year = next_cell.year, month = next_cell.month, day = next_cell.day)
            ### comparison error case
            if(next_date < current_date):
                if(unordered_date_error == 0):
                    unordered_date_error = 1
                    print("Unordered dates:", current_date, "and", next_date)
                    if (in_range == 1):
                        error_ws.cell(row = COUNT, column = 1).value = file 
                        error_ws.cell(row = COUNT, column = 2).value = "Dates not in order"
                        increment()
                        error_wb.save(dir_error)
                i=j
                j+=1
                continue
            ### must continue until we reach m
            if(next_date <= start_of_m):
                minimum = ws.cell(row = j, column = 7).value
                if(issued in ws.cell(row = j, column = 2).value.lower()):
                    if(next_date >= (start_of_m - timedelta(days = 90))):
                        print("issued")
                        ap_issued = 1
                elif((cancelled in ws.cell(row = j, column = 2).value.lower()) | (purchased in ws.cell(row = j, column = 2).value.lower())):
                    print("cancelled")
                    ap_issued = 0
                i=j
                j+=1
                continue
            elif((next_date > start_of_m) & (next_date <= end_of_m)):
                in_range = 1
                if(minimum > ws.cell(row = j, column = 7).value):
                    minimum = ws.cell(row = j, column = 7).value
                    i = j
                if(issued in ws.cell(row = j, column = 2).value.lower()):
                    print("ap issued")
                    ap_issued = 1
                    return minimum, 0, i, 0, ap_issued
                elif((cancelled in ws.cell(row = j, column = 2).value.lower()) | (purchased in ws.cell(row = j, column = 2).value.lower())):
                    print("ap_cancelled")
                    ap_issued = 0
                    return 0, 0, i, 0, ap_issued
                if(ws.cell(row = j+1, column = 7).value == None):
                    return minimum, 0, i, 0, ap_issued
                j+=1
                continue
            ### outside of range, take last minimum
            elif(next_date > end_of_m):
                # return{'partial_roi' : minimum, 'error' : 0, 'r' : i, 'EoF' : 0}
                return minimum, 0, i, 0, ap_issued
        except AttributeError:
            if(ws.cell(row = j+1, column = 7).value != None):
                if(missing_date_error == 0):
                    missing_date_error = 1
                    print("Missing Date")
                    if(in_range == 1):
                        error_ws.cell(row = COUNT, column = 1).value = file 
                        error_ws.cell(row = COUNT, column = 2).value = "Missing Dates"
                        increment()
                        error_wb.save(dir_error)
                        return minimum, 1, i, 0, ap_issued
                j+=1
                continue
            elif(ws.cell(row = j+2, column = 7).value != None):
                if(missing_date_error == 0):
                    missing_date_error = 1
                    print("Missing Date")
                    if(in_range == 1):
                        error_ws.cell(row = COUNT, column = 1).value = file 
                        error_ws.cell(row = COUNT, column = 2).value = "Missing Dates"
                        increment()
                        error_wb.save(dir_error)
                        return minimum, 1, i, 0, ap_issued
                j+=2
                continue
            ### missing dates should be somewhat frequent in the beginning
            else:
                return minimum, 0, i, 1, ap_issued
        except TypeError:
            error_ws.cell(row = COUNT, column = 1).value = file 
            error_ws.cell(row = COUNT, column = 2).value = "Other Error"
            increment()
            error_wb.save(dir_error)
            return 0,1,i,0,ap_issued
        except:
            print("Unexpected error")
            error_ws.cell(row = COUNT, column = 1).value = file 
            error_ws.cell(row = COUNT, column = 2).value = "Unexpected error"
            increment()
            error_wb.save(dir_error)
            return 0,1,i,0,ap_issued
    
    print("Done partial roi")
    # return{'partial_roi' : minimum, 'error' : 0, 'r' : i, 'EoF' : 0}
    return minimum, 0, i, 0, ap_issued

def write_principal_column(file, p1, p2, p3, avg, roi):
    principal_ws.cell(row = p_COUNT, column = 1).value = file
    principal_ws.cell(row = p_COUNT, column = 2).value = p1
    principal_ws.cell(row = p_COUNT, column = 3).value = p2
    principal_ws.cell(row = p_COUNT, column = 4).value = p3
    principal_ws.cell(row = p_COUNT, column = 5).value = avg
    principal_ws.cell(row = p_COUNT, column = 6).value = roi
    principal_wb.save(dir_principal)
def write_output_file(file, file_out, roi, r):
    book = load_workbook(file)
    ws = book.active
    cell = ws.cell(row=r,column=1).value
    current_date = date(year = cell.year, month = cell.month, day = cell.day)
    origin = 'G' + str(10)
    formula = "=G9-C10+D10+E10+F10"
    thin = Side(border_style="thin", color="000000")
    if(current_date >= end_of_quarter):
        ws.insert_rows(r)
    else:
        r+=1
        ws.insert_rows(r)
    j = 1
    while(j < 8):
        cell = ws.cell(row = 9, column = j)
        new_cell = ws.cell(row=r, column = j)
        if cell.has_style:
            new_cell._style = copy(cell._style)
        else:
            print("No style detected")
            ws.cell(row = r-1, column = j).border = Border(top = thin, right = thin, left = thin, bottom = thin)        
        new_cell.font = Font(bold = None)
        j +=1
    ws.cell(row = r, column = 1).value = end_of_quarter
    ws.cell(row = r, column = 2).value = roi_string
    ws.cell(row = r, column = 5).value = roi
    ws.cell(row = r, column = 5).alignment = Alignment(horizontal = 'right')
    fcell = 'G' + str(r)
    gcell = 'G' + str(r+1)
    ws[fcell].value = Translator(formula, origin).translate_formula(fcell)
    if(ws[gcell].value != None):
        ws[gcell].value = Translator(formula, origin).translate_formula(gcell)
    print("End of file")
    book.save(file_out)
    book.close()
    
def process_file(file_in, file_out, file):
    out_of_range = 0
    r = 9
    p1 = 0
    p2 = 0
    p3 = 0
    ap = 0
    os.chdir(dir_pre)
    book = load_workbook(file_in, data_only = True)
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
        p1, error, r, EoF, ap = partial_roi(ws, file, start_of_m1, end_of_m1, r, 0)
        p1 = round(p1, 2)
        if(error):
            print("error: check log file")
            return
        if(ap):
            p1 = 0
        if(EoF):
            p2 = p1
            p3 = p1
        else:
            print("passing in row", r, " to p2")
            p2, error, r, EoF, ap = partial_roi(ws, file, start_of_m2, end_of_m2, r, ap)
            p2 = round(p2, 2)
            if(error):
                print("error: check log file")
                return
        if(ap):
            p2 = 0
        if(EoF):
            p3 = p2
        else:
            p3, error, r, EoF, ap = partial_roi(ws, file, start_of_m3, end_of_m3, r, ap)
            p3 = round(p3, 2)
            if(error):
                print("error: check log file")
                return
        if(ap):
            p3 = 0     
        print("P1:",p1)
        print("P2:",p2)
        print("P3:",p3)
        avg = (p1 + p2 + p3)/3
        roi = (p1 + p2 + p3) * declared_roi/(3*100)
        roi = round(roi, 2)
        if(write_principal == True):
            write_principal_column(file, p1, p2, p3, avg, roi)
            p_increment()
        if(write_statements == True):
            write_output_file(file, file_out, roi, r)
       
    except AttributeError:
       print("Missing date in first row")
       error_ws.cell(row = COUNT, column = 1).value = file 
       error_ws.cell(row = COUNT, column = 2).value = "Missing first row"
       increment()
       error_wb.save(dir_error)
       book.close()
       return
    except TypeError:
       print("Possible incorrect value")
       error_ws.cell(row = COUNT, column = 1).value = file 
       error_ws.cell(row = COUNT, column = 2).value = "Incorrect value"
       increment()
       error_wb.save(dir_error)
       book.close()
       return
    except:
       print("Unexpected error")
       error_ws.cell(row = COUNT, column = 1).value = file 
       error_ws.cell(row = COUNT, column = 2).value = "Unexpected error"
       increment()
       error_wb.save(dir_error)
       book.close() 
       return
    book.close()
if(write_principal == True):
    principal_ws.cell(row =  1, column = 1).value = "File No"
    principal_ws.cell(row =  1, column = 2).value = "P1"
    principal_ws.cell(row =  1, column = 3).value = "P2"
    principal_ws.cell(row =  1, column = 4).value = "P3"
    principal_ws.cell(row =  1, column = 5).value = "Avg"
    principal_ws.cell(row =  1, column = 6).value = "ROI"
    p_increment()
for file in sorted( filter(lambda x: not (x.startswith('~') or x.startswith('.')), dir_list) ):
    file_in = os.path.join(dir_pre, file)
    file_out = os.path.join(dir_post, file)
    #file_in = file_in.replace(os.sep, '/')
    print(file_in)
    process_file(file_in, file_out, file)   
            
    
        
        
        























