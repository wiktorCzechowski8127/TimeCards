# -*- coding: utf-8 -*-

from excel import *
from  employees import *
import datetime
import os
import time

#from mfrc522 import SimpleMFRC522
#reader = SimpleMFRC522()
COOLDOWN_TIME_IN_SEC = 2

def changeFile(employees, exc, now, currentExcelFile):   
    excToCopy = excelSheetC(currentExcelFile)
    excToCopy.create_workbook()
    excToCopy.generateExcelTemplate(employees, now)
    excToCopy.saveSheet()
    excToCopy.lastModificationTime = time.ctime(os.path.getmtime(excToCopy.file))
    return excToCopy

def defineMonthAndFile():
    now =  datetime.datetime.now()
    month = now.month
    file = str(now.month) + "_" + str(now.year) + ".xlsx"
    return(file, month)

#employees, exc, now, currentMonth, currentExcelFile, lastModificationTime)
def checkMonthAndFile(employees, exc, now, currentMonth):
    isSafeMode = False
    if (now.month != currentMonth):
        print("Changing month")
        isSafeMode = True
        employeesAfterUpdate = employeesC()
        if(employeesAfterUpdate.updateEmployeesStatus(exc)):
            employeesAfterUpdate.addNewEmployees(exc, now)
            employeesAfterUpdate.deleteLayedEmployees()
            employees = employeesAfterUpdate 

            currentExcelFile = str(now.month) + "_" + str(now.year) + ".xlsx"
            exc = changeFile(employees, exc, now, currentExcelFile)
            currentMonth = now.month
            isSafeMode = False
    return employees, exc, currentMonth, isSafeMode

def updateEmployees(employees, exc, now):
    print("Updating Employees")
    isSafeMode = True
    employeesAfterUpdate = employeesC()
    if(employeesAfterUpdate.updateEmployeesStatus(exc)):
        employeesAfterUpdate.addNewEmployees(exc, now)
        employees = employeesAfterUpdate
        isSafeMode = False
    exc.saveSheet()
    exc.lastModificationTime = time.ctime(os.path.getmtime(exc.file))
    return employees, exc, isSafeMode



if __name__=="__main__": 
    
    #currentExcelFile = str(now.month) + "_" + str(now.year) + ".xlsx"
    #currentExcelFile, currentMonth = defineMonthAndFile()
    isSafeMode = False
    #currentExcelFile = '2_2024.xlsx'
    NOWYEAR = 2024
    NOWMONTH = 2
    NOWDAY = 1
    NOWHOUR = 1
    NOWMINUTES = 0
    NOWSECONDS = 0
    now = datetime.datetime(NOWYEAR, NOWMONTH, NOWDAY, NOWHOUR, NOWMINUTES, NOWSECONDS)
    lastMinute = 0
    currentMonth = now.month
    exc = excelSheetC(str(now.month) + "_" + str(now.year) + ".xlsx")
    employees = employeesC()
    #employees.readEmployeesFromXlsx(exc)

    #while

    i = 0
    #updateEmployees(exc, now)
    lastReadedId = 0
    lastReadedTimeStamp = 0
    coolDown = 0

    while(True):
        employees, exc, currentMonth, isSafeMode = \
            checkMonthAndFile(employees, exc, now, currentMonth)
        
        modificationTimeSample = time.ctime(os.path.getmtime(exc.file))
        if (exc.lastModificationTime != modificationTimeSample):
            employees, exc, isSafeMode = updateEmployees(employees, exc, now)

        if (isSafeMode):
            print("Safe mode enabled")
        else:
            readedTokenId = None #int(reader.read_no_block()[0])
            if ((readedTokenId != None) and (id != lastReadedId)):
                lastReadedTimeStamp = time.time()
                lastReadedId = readedTokenId
                employeeId = employees.checkReadedId(readedTokenId)
                if (employeeId != None):
                    currentTime = now #datetime.datetime.now()
                    exc.inputTimestampIntoExcel(readedTokenId, employeeId, currentTime)
                else:
                    exc.inputTokenIdToExce(readedTokenId)
                print("LastReadedId: %s recieved %.2f seconds ago" % (hex(lastReadedId),coolDown))
            coolDown = abs(time.time() - lastReadedTimeStamp)
            if (coolDown > COOLDOWN_TIME_IN_SEC):
                lastReadedId = 0


        #To Delete   
        now2 = datetime.datetime.now()
        if (NOWMINUTES == 60):
            NOWMINUTES = 0
            NOWHOUR += 1
            if (NOWHOUR == 24):
                NOWHOUR = 1
                NOWDAY += 1
        if(NOWDAY > 27):
            NOWDAY = 1
            NOWMONTH += 1
        
            if(NOWMONTH > 12):
                NOWMONTH = 1
                NOWYEAR +=1
        now = datetime.datetime(NOWYEAR, NOWMONTH, NOWDAY, NOWHOUR, NOWMINUTES, NOWSECONDS)
        
        if(now2.second != lastMinute):
            print(now)
            lastMinute = now2.second
            NOWMINUTES += 1
