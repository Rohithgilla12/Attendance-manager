from datetime import datetime
import openpyxl
import os
from shutil import copyfile
import string
test=input("Enter the list containing roll numbers :")
# test=[25,32]
currentDay = datetime.now().day
currentMonth = datetime.now().month
currentYear = datetime.now().year
filename=str(currentMonth)+"_"+str(currentYear)+".xlsx"
def leap(y) :
    if (y%4):
        if (y%100):
             if (y%400):
                 return True
        else :
            return True
    else :
        return False
def attendence(test):
    temp=str(string.letters)
    # print temp
    temp=temp.split('z')[1]
    # print temp
    temp2=['AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ']
    cells=list(temp)
    cells+=temp2
    m31=[1,3,5,7,8,10,12]
    m30=[4,6,9,11]
    m28=[2]
    i=int(currentMonth)
    if i in m31 :
        n=31
    elif i in m30:
        n=30
    elif((i in m28) and (not(leap(int(currentYear))))) :
        n=28
    else :
        n=29
    try:
        op=open(filename,'r')
        op.close()
    except IOError:
        src=os.getcwd()+"/default.xlsx"
        dst=os.getcwd()+"/"+filename
        copyfile(src,dst)
        wb=openpyxl.load_workbook(filename)
        sheet=wb.get_sheet_by_name('Sheet1')
        i=1
        while i<=n:
            sheet[str(cells[i+1])+str(2)]=str(i)+"/"+str(currentMonth)+"/"+str(currentYear)
            i+=1
        wb.save(filename)
    wb=openpyxl.load_workbook(filename)
    sheet=wb.get_sheet_by_name('Sheet1')
    for i in test:

        # print len(cells)
        # print str(cells[int(currentDay)+1])+str(i+2)
        sheet[str(cells[currentDay+1])+str(i+2)]=1
    wb.save(filename)
attendence(test)
