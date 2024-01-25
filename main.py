# -*- coding: utf-8 -*-

from excel import *
from  employees import *
import datetime
import os
import time


def changeFile(exc, now, currentExcelFile, lastModificationTime):   
    excToCopy = excelSheetC(currentExcelFile)
    excToCopy.create_workbook()
    employees = employeesC()
    employees.updateEmployeesStatus(exc)
    employees.addNewEmployees(exc, now)
    employees.deleteLayedEmployees()
    exc.saveSheet()
    excToCopy.generateExcelTemplate(employees, now)
    excToCopy.saveSheet()
    return excToCopy

def defineMonthAndFile():
    now =  datetime.datetime.now()
    month = now.month
    file = str(now.month) + "_" + str(now.year) + ".xlsx"
    return(file, month)
    
def checkMonthAndFile(currentMonth, exc, now, currentExcelFile, lastModificationTime):
    #now =  datetime.datetime.now()
    #now = datetime.datetime(2024, 1, 1)
    if (now.month != currentMonth):
        currentExcelFile = str(now.month) + "_" + str(now.year) + ".xlsx"
        exc = changeFile(exc, now, currentExcelFile, lastModificationTime)
        currentMonth = now.month
        ti_m = os.path.getmtime(currentExcelFile)
        lastModificationTime = time.ctime(ti_m)
    return exc, currentExcelFile, lastModificationTime, currentMonth

def updateEmployees(exc, now):
    employees = employeesC()
    if(employees.updateEmployeesStatus(exc)):
    #employees.deleteLayedEmployees()
    #now = datetime.datetime(2024, 2, 1) #TO DO
        employees.addNewEmployees(exc, now)
    exc.saveSheet()



if __name__=="__main__": 

    
    #currentExcelFile = str(now.month) + "_" + str(now.year) + ".xlsx"
    #currentExcelFile, currentMonth = defineMonthAndFile()
    isSafeMode = False
    lastModificationTime = ''
    currentExcelFile = '2_2024.xlsx'
    currentMonth = 2
    exc = excelSheetC(currentExcelFile)
    #employees = employeesC()
    #employees.readEmployeesFromXlsx(exc)

    #while
    now = datetime.datetime(2024, 2, 1)
    i = 0
    #updateEmployees(exc, now)

    while(True):
        exc, currentExcelFile, lastModificationTime, currentMonth = \
            checkMonthAndFile(currentMonth, exc, now, currentExcelFile, lastModificationTime)
        modificationTimeSample = time.ctime(os.path.getmtime(currentExcelFile))
        if (lastModificationTime != modificationTimeSample):
            updateEmployees(exc, now)
            lastModificationTime = time.ctime(os.path.getmtime(currentExcelFile))
            i += 1
            print(i)
        
        if(i == 2):
            print('month changing')
            now = datetime.datetime(2024, currentMonth + 1, 1)
            i = 0
            print(now)


#add safe mode



    '''
    employees = employeesC()
    employees.readEmployeesFromYaml()
    if(employees.isClassFilledCorectly):

        employees.showEmployees()

        now =  datetime.datetime.now()
        currentExcelFile = str(now.month) + "_" + str(now.year) + ".xlsx"
        timeSheet = []
        for i in range (30):
            timeSheet.append(excelSheetC(currentExcelFile))
            #timeSheet[i].readEmployeesFromYaml()
            timeSheet[i].checkFileAndInitialize()
            timeSheet[i].generateExcelTemplate(employees)
            timeSheet[i].saveSheet()
            timeSheet[i].sheet.cell(row=23, column=2).value = 3
            timeSheet[i].saveSheet()
            print('dupa')
            timeSheet[i].initializeSheet()
            timeSheet[i].sheet.cell(row=23, column=3).value = 4
            timeSheet[i].saveSheet()

        
        
    else:
        print("Class not filled correclty")
        

    '''  
    
    
    
    #employees.addEmployee()
    #employees.showEmployees()
    #employees.changeEmployeeName()
    #employees.changeEmployeeTokenId()
    #employees.showEmployees()
    #employees.writeEmployeesToYaml()
    #print("Program stop running:", int(not(employees.isClassFilledCorectly)))
            

    #employeeList[i] = employee
   
    
   
#employee.printclass(employeeList[1])
#print(employeeList[1])
#print(tmp['employee(0)'])

#e1 = employee(0, 2137, 'marcin')
#print(e1.employeeid)


        
        
        
    



