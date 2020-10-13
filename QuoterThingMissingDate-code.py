# -*- coding: utf-8 -*-
"""
Created on Wed Sep  4 10:16:38 2019

@author: stefflc
@authorfordateupdates: doshig
"""


'''
Take excel File
1st Tab is something
2nd tab needs sums from 1st TabError
E11 is upper left
K11 is upper right

'''

### TO RUN THIS SCRIPT ON MISSING DATE FILES, python EXECUTE IN TERMINAL AND THEN SELECT THE FILE ###

##Might want to create a backup folder/file every time this program is run against a file?

import xlrd
# from xlrd.sheet import ctype_text
from tkinter import *
from tkinter.filedialog import askopenfilename ##tkinter
# import pyodbc
# import sqlite3
import datetime
# import sys
import os
import xlsxwriter
import copy

#print(sys.path)


## Update 1/21/2020
## Add logic to exclude snrqty = conqty

## Update 01/28/2020
## Use schqty not conqty

## Update 04/21/2020
## Skip over empty date field that are used for RnD products

###All this is needed to popup the window then close it again###
root = Tk()
root.withdraw()
root.update()
fname = askopenfilename()
root.destroy()

# fname = "PW_ZSACNXLS_20200428_100704" ##For testing only, static file
print("filename: ", fname)
###




try:
    wb = xlrd.open_workbook(fname)
except FileNotFoundError:
    print("error")
    

sheet_names = wb.sheet_names()
#print("how many sheets?: ", len(sheet_names))
print("how many sheets?: ", wb.nsheets) #Use built in methods from xlrd
NUM_SHEETS = wb.nsheets

print("Looking at 1st tab: ", sheet_names[0])
MAIN_SHEET = sheet_names[0] #Grab 1st sheet/tab
for y in sheet_names:
    print("all sheet name(s): ", y)



#for x in range(NUM_SHEETS):
#    print("sheet position: ", x)
#    print("sheet name: ", sheet_names[x])
##    print("sheet name: ", wb.sheet_by_index(x))
##    print("sheet name: ", wb.sheet_by_name(sheet_names[x]))  #Sheet objects
    
HIT_SH = wb.sheet_by_name(MAIN_SHEET)

print("how many columns in HIT LIST?: ", HIT_SH.ncols)
print("how many rows in HIT LIST?: ", HIT_SH.nrows)

COLUMN_LOCATION = list()

#for n in range(0, HIT_SH.nrows):


###Column constants###
PRODUCT = 4
DEL_DATE = 5
SCH_QTY = 6
SNR_QTY = 7
CON_DATE = 8
CON_QTY = 9
UOM = 10






###Positional constants ###
ROW_11 = 10
COL_A = 0
COL_E = 4

tuple_list = list()

small_dict = {}


'''Get header information'''
ROW_1 = 0
ROW_2 = 1
ROW_3 = 2
ROW_4 = 3
ROW_5 = 4
ROW_6 = 5
ROW_7 = 6

COL_2 = 2
interfaceType = HIT_SH.col_values(COL_2)[ROW_1]
ownerPartner = HIT_SH.col_values(COL_2)[ROW_2]
partner = HIT_SH.col_values(COL_2)[ROW_3]
selectionProfile = HIT_SH.col_values(COL_2)[ROW_4]
selectionProfile2 = HIT_SH.col_values(COL_2)[ROW_5]
createdBy = HIT_SH.col_values(COL_2)[ROW_6]
createdOn = HIT_SH.col_values(COL_2)[ROW_7]


for row in range(ROW_11, HIT_SH.nrows): ###exclude the header row
# for row in range(ROW_11, 20): ###exclude the header row

    # for col in range(COL_E, HIT_SH.ncols):
        # print("col: ",col, row, HIT_SH.col_values(col)[row])
    


    dateCreated = datetime.datetime.now()
    dateModified = datetime.datetime.now()
    
    product = HIT_SH.col_values(PRODUCT)[row]
    deldate = HIT_SH.col_values(DEL_DATE)[row]
    schqty = HIT_SH.col_values(SCH_QTY)[row]
    snrqty = HIT_SH.col_values(SNR_QTY)[row]
    #if snrqty == schqty -> exclude
    condate = HIT_SH.col_values(CON_DATE)[row]
    conqty = HIT_SH.col_values(CON_QTY)[row]
    uom = HIT_SH.col_values(UOM)[row]

    fred = datetime.datetime(*xlrd.xldate_as_tuple(deldate, wb.datemode))
    fred = fred.date() ##dueDate is a datefield not datetime field
    # print("fred: ", fred)
    deldate = fred

    #IF STATEMENT ADDED BELOW TO SKIP CONDATE CELL IF ITS BLANK(COL 8)
    if condate == "":
        print("empty")
    else:
        fred = datetime.datetime(*xlrd.xldate_as_tuple(condate, wb.datemode))
        fred = fred.date()
        # print("fred2: ", fred)
        condate = fred

        condate = condate.strftime("%m/%d/%Y")

        '''
        These are excluded per PWA rep 01/21/2020
        '''
        if(schqty == snrqty):
            print("skipping row %s for: %s. Reason: SCH Qty %s = SNR Qty %s."%(row,product,schqty,snrqty))
            pass
        else:
            t = (product, condate)
            if t in tuple_list:
                # small_dict[t] = small_dict[t] + conqty
                small_dict[t] = small_dict[t] + schqty
            else:
                tuple_list.append(t)
                # small_dict[t] = conqty
                small_dict[t] = schqty

            # print(small_dict, '\n')


e = datetime.datetime.now()
e = str(e.month)+"-"+str(e.day)+"-"+str(e.year)+"--"+str(e.hour)+"h-"+str(e.minute)+"m"

titleexport = 'PWA '+str(e)+'.xlsx'

try :
    desktoppath = os.path.join(os.path.expanduser("~"), "Desktop")
except Exception as ex:
    print(ex)


try:
    folder = str(desktoppath)+"\\PWA"
    if not os.path.exists(folder):
        os.makedirs(folder)
except Exception as ex:
    print(ex)




try:
    way = str(desktoppath)+"\\PWA\\"+str(titleexport)

    workbook = xlsxwriter.Workbook(way)
except :
    print("desktoppatherror")
    

tab = workbook.add_worksheet("output")
instructions = workbook.add_worksheet("instructions")


# interfaceType = HIT_SH.col_values(4)[ROW_1]
# ownerPartner = HIT_SH.col_values(4)[ROW_2]
# partner = HIT_SH.col_values(4)[ROW_3]
# selectionProfile = HIT_SH.col_values(4)[ROW_4]
# selectionProfile2 = HIT_SH.col_values(4)[ROW_5]
# createdBy = HIT_SH.col_values(4)[ROW_6]
# createdOn

tab.write(0, 0, "Interface Type")
tab.write(0, 1, interfaceType)

tab.write(1, 0, "Onwer Partner")
tab.write(1, 1, ownerPartner)

tab.write(2, 0, "Partner: ")
tab.write(2, 1, partner)

tab.write(3, 0, "Selection Profile")
tab.write(3, 1, selectionProfile)

tab.write(4, 0, "Selection Profile")
tab.write(4, 1, selectionProfile2)

tab.write(5, 0, "Created by: ")
tab.write(5, 1, createdBy)

tab.write(6, 0, "Created on: ")
tab.write(6, 1, createdOn)


tab.write(10, 0, "OUTPUT CREATED: "+ dateCreated.strftime("%m/%d/%Y"))

tab.write(11, 0, "PART NUMBER")
tab.write(11, 1, "CON. DATE")
tab.write(11, 2, "TOTAL CON. QTY")

copy_dict = copy.deepcopy(small_dict)



ROW_OFFSET = 12
for index, value in enumerate(small_dict.items()):
    # (row, column, text)
    # print("aaaa: ", index)
    # print("baaa: ", value)
    pn = value[0][0]
    date = value[0][1]
    qty = value[1]
    # print("pn: ", pn)
    # print("date: ", date)
    # print("qty: ", qty)
    tab.write(index+ROW_OFFSET, 0, pn)
    tab.write(index+ROW_OFFSET, 1, date)
    tab.write(index+ROW_OFFSET, 2, qty)


instructions.write(0,0, "INSTRUCTIONS FOR PASTING")
instructions.write(1,0, "1. TO GET CORRECT PASTE ORDER ON PWA SHEET")
instructions.write(2,0, "2. CLICK FILTER ARROW OF PRODUCT (Cell A1)")
instructions.write(3,0, "3. SORT BY COLOR - > CUSTOM SORT")

instructions.write(4,0, "4. under column -> sort by : Product")
instructions.write(5,0, "5. Click OK at bottom")

instructions.write(6,0, "6. Important! Choose 'sort anything that looks like a number, as a number' (TOP option)")
workbook.close()


exit







