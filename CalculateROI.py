import os
import calendar
import datetime
from copy import copy
from datetime import date
from openpyxl.formula.translate import Translator
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment

declared_roi = 1.5
issued = "AP Letter Issued"
cancelled = "AP Letter Cancelled"
purchased = "Home Purchased"
#######################
### Get directories ###
#######################
dir = os.getcwd()
dir_pre = os.path.join(dir, "Q4Before")
dir_post = os.path.join(dir, "Q4After")
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
Quarter = ((m-1)//3 + 1)
pstr = str(year) + "Q" + str(Quarter) + "principal_column.xlsx"
dir_principal = os.path.join(dir, pstr)
roi_string = "ROI " + str(input_day.year) + "Q" + str(Quarter) + ": " + "{:.2f}%".format(declared_roi)
print(roi_string)
if (Quarter == 1):
    month1 = 1
    month2 = 2
    month3 = 3
elif (Quarter == 2):
    month1 = 4
    month2 = 5
    month3 = 6
elif (Quarter == 3):
    month1 = 7
    month2 = 8
    month3 = 9
else:
    month1 = 10
    month2 = 11
    month3 = 12
    
end_of_p1 = date(year = input_day.year, month = month1, day = 1)
end_of_p2 = date(year = input_day.year, month = month2, day = 1)
end_of_p3 = date(year = input_day.year, month = month3, day = 1)

lastday = calendar.monthrange(input_day.year, month3)[1]
end_of_quarter = date(year = input_day.year, month = month3, day = lastday)


print("Month 1:", calendar.month_name[month1], "Last day:", calendar.monthrange(input_day.year, month1)[1])
print("Month 2:", calendar.month_name[month2], "Last day:", calendar.monthrange(input_day.year, month2)[1])
print("Month 3:", calendar.month_name[month3], "Last day:", calendar.monthrange(input_day.year, month3)[1])    
print("End of Quarter:", calendar.month_name[month3], lastday)
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
   
    cell = ws.cell(row=r,column=1).value
    if ((cell == None) | (cell == "Date")):
        print("First row mismatch.")
        error_ws.cell(row = COUNT, column = 1).value = file 
        error_ws.cell(row = COUNT, column = 2).value = "Row mismatch"
        increment()
        error_wb.save(dir_error)
        return
    current_date = date(year = cell.year, month = cell.month, day = cell.day)
    first_date = date(year = cell.year, month = cell.month, day = cell.day)
    if (first_date > end_of_quarter):
            out_of_range = 1
    if(out_of_range):
        print("First date is after the current quarter. Continuing...")
        error_ws.cell(row = COUNT, column = 1).value = file 
        error_ws.cell(row = COUNT, column = 3).value = "After quarter"
        increment()
        error_wb.save(dir_error)
        return
    while (current_date < end_of_quarter):
        cell = ws.cell(row=r,column=1).value
        if ~hasattr(cell, 'year'):    
            if((ws.cell(row = r, column = 7).value != None)|(ws.cell(row = r+1, column = 7).value != None)):
               r+=1
               continue
        print(cell)
        # if(((cell == None) | (cell == " ")) & (ws.cell(row = r, column = 7).value != None)):
        #     print("Empty Rows require correction")
        #     error_ws.cell(row = COUNT, column = 1).value = file 
        #     error_ws.cell(row = COUNT, column = 2).value = "Empty rows/missing dates detected. First instance is row " + str(r) 
        #     increment()
        #     error_wb.save(dir_error)
        if(cell != None):
            if(ws.cell(row = r, column = 7).value == None):
                print("Missing Balance")
                error_ws.cell(row = COUNT, column = 1).value = file 
                error_ws.cell(row = COUNT, column = 2).value = "Missing Balance on row " + str(r)
                increment()
                error_wb.save(dir_error)
                return
            if (r > 9):
                if(ws.cell(row = r-1, column = 2).value == issued):
                    ap_issued = 1
                    ap_cancelled = 0
                    print(issued)
                elif(ws.cell(row = r-1, column = 2).value == cancelled): 
                    ap_cancelled = 1
                    ap_issued = 0
                    print(cancelled)
                elif(ws.cell(row = r, column = 2).value == purchased):
                    ap_cancelled = 1
                    ap_issued = 0
                    print(purchased)
                # if ((cell.year < current_date.year) | ((cell.year == current_date.year) & (cell.month < current_date.month)) | ((cell.year == current_date.year) & (cell.month == current_date.month) & (cell.day < current_date.day))):    
                #     print("Dates are not in ascending order")
                #     error_ws.cell(row = COUNT, column = 1).value = file 
                #     error_ws.cell(row = COUNT, column = 2).value = "Dates are not in ascending order. First instance is row " + str(r) 
                #     increment()
                #     error_wb.save(dir_error)
                #     return
            current_date = date(year = cell.year, month = cell.month, day = cell.day)
            print(current_date)
            if (current_date > end_of_p1):
                if (r > 9):
                    p1 = ws.cell(row = r-1, column = 7).value
                else:
                    p1 = 0
                break
            r += 1
        else:
            p1 = ws.cell(row = r-1, column = 7).value
            break
    if(ap_issued):
        p1 = 0
    while (current_date < end_of_quarter):
        cell = ws.cell(row=r,column=1).value
        if ~hasattr(cell, 'year'):    
            if((ws.cell(row = r, column = 7).value != None)|(ws.cell(row = r+1, column = 7).value != None)):
               r+=1
               continue
        if(r>9):
            if(ws.cell(row = r-1, column = 2).value == issued):
                ap_issued = 1
                ap_cancelled = 0
                print(issued)
            elif(ws.cell(row = r-1, column = 2).value == cancelled): 
                ap_cancelled = 1
                ap_issued = 0
                print(cancelled)
            elif(ws.cell(row = r, column = 2).value == purchased):
                ap_cancelled = 1
                ap_issued = 0
                print(purchased)
        if (cell != None):
            if(ws.cell(row = r, column = 7).value == None):
                print("Missing Balance")
                error_ws.cell(row = COUNT, column = 1).value = file 
                error_ws.cell(row = COUNT, column = 2).value = "Missing Balance on row " + str(r)
                increment()
                error_wb.save(dir_error)
                return
            current_date = date(year = cell.year, month = cell.month, day = cell.day)
            if (current_date > end_of_p2):
                if (r > 9):
                    p2 = ws.cell(row = r-1, column = 7).value
                else:
                    p2 = 0
                break
            r += 1
        else:
            p2 = ws.cell(row = r-1, column = 7).value
            break
    if(ap_issued):
        p2 = 0
    while (current_date < end_of_quarter):
        cell = ws.cell(row=r,column=1).value
        if ~hasattr(cell, 'year'):    
            if((ws.cell(row = r, column = 7).value != None)|(ws.cell(row = r+1, column = 7).value != None)):
                r+=1
                continue
        if(r>9):
            if(ws.cell(row = r-1, column = 2).value == issued):
                ap_issued = 1
                ap_cancelled = 0
                print(issued)
            elif(ws.cell(row = r-1, column = 2).value == cancelled): 
                ap_cancelled = 1
                ap_issued = 0
                print(cancelled)
            elif(ws.cell(row = r, column = 2).value == purchased):
                ap_cancelled = 1
                ap_issued = 0
                print(purchased)
        if (cell != None):
            if(ws.cell(row = r, column = 7).value == None):
                print("Missing Balance")
                error_ws.cell(row = COUNT, column = 1).value = file 
                error_ws.cell(row = COUNT, column = 2).value = "Missing Balance on row " + str(r)
                increment()
                error_wb.save(dir_error)
                return
            current_date = date(year = cell.year, month = cell.month, day = cell.day)
            if (current_date > end_of_p3):
                p3 = ws.cell(row = r-1, column = 7).value
                r += 1
                break
            r += 1
        else:
            p3 = ws.cell(row = r-1, column = 7).value
            break
    if(ap_issued):
        p3 = 0
        
    print("P1:",p1)
    print("P2:",p2)
    print("P3:",p3)
    roi = (p1 + p2 + p3) * declared_roi/(3*100)
    print("ROI:",roi) 
    write_principal_column(file, p1, p2, p3, roi)
    p_increment()
    #roi = (p1 + p2 + p3) * declared_roi/(3*100)
    #print("ROI:",roi)
    
    book.close()
    # book = load_workbook(file)
    # ws = book.active
    # cell = ws.cell(row=r-1,column=1).value
    # current_date = date(year = cell.year, month = cell.month, day = cell.day)
    # origin = 'G' + str(10)
    # formula = "=G9-C10+D10+E10+F10"
    # thin = Side(border_style="thin", color="000000")
    # if(current_date >= end_of_quarter):
    #     ws.insert_rows(r-1)
    # else:
        # r+=1
    #     ws.insert_rows(r-1)
    # j = 1
    # while(j < 8):
    #     cell = ws.cell(row = 9, column = j)
    #     new_cell = ws.cell(row=r-1, column = j)
    #     if cell.has_style:
    #         new_cell._style = copy(cell._style)
    #     else:
    #         print("No style detected")
    #         ws.cell(row = r-1, column = j).border = Border(top = thin, right = thin, left = thin, bottom = thin)        
    #     new_cell.font = Font(bold = None)
    #     j +=1
    # ws.cell(row = r-1, column = 1).value = end_of_quarter
    # ws.cell(row = r-1, column = 2).value = roi_string
    # ws.cell(row = r-1, column = 5).value = roi
    # ws.cell(row = r-1, column = 5).alignment = Alignment(horizontal = 'right')
    
    # fcell = 'G' + str(r-1)
    # gcell = 'G' + str(r)
    # ws[fcell].value = Translator(formula, origin).translate_formula(fcell)
    # if(ws[gcell].value != None):
    #     ws[gcell].value = Translator(formula, origin).translate_formula(gcell)
    # print("End of file")
    # book.save(file_out)
    # book.close()
  

for file in sorted( filter(lambda x: not (x.startswith('~') or x.startswith('.')), dir_list) ):
#file = "0733C2.xlsx"
    file_in = os.path.join(dir_pre, file)
    file_out = os.path.join(dir_post, file)
    #file_in = file_in.replace(os.sep, '/')
    print(file_in)
    process_file(file_in, file_out, file)
