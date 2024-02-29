# -*- coding: utf-8 -*-

from excel import *
from  employees import *
import datetime
import os
import time


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
    isSafeMode = True
    employeesAfterUpdate = employeesC()
    if(employeesAfterUpdate.updateEmployeesStatus(exc)):
    #employees.deleteLayedEmployees()
    #now = datetime.datetime(2024, 2, 1) #TO DO
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
    now = datetime.datetime(2024, 2, 1)
    currentMonth = now.month
    exc = excelSheetC(str(now.month) + "_" + str(now.year) + ".xlsx")
    employees = employeesC()
    #employees.readEmployeesFromXlsx(exc)

    #while

    i = 0
    #updateEmployees(exc, now)
    lastReadedId = 0
    lastReadedTimeStamp = 0

    while(True):
        employees, exc, currentMonth, isSafeMode = \
            checkMonthAndFile(employees, exc, now, currentMonth)
        
        modificationTimeSample = time.ctime(os.path.getmtime(exc.file))
        if (exc.lastModificationTime != modificationTimeSample):
            employees, exc, isSafeMode = updateEmployees(employees, exc, now)
            i += 1
            print(i)

        if (isSafeMode):
            print("Safe mode enabled")
        else:
            print("RFID running")
            id = 501 #reader.read_no_block()[0]
            if ((id != None) and (id != lastReadedId)):
                lastReadedTimeStamp = time.time()
                lastReadedId = id
                print("ID: %s" %(id))
                employeeId = employees.checkReadedId(id)
                if (employeeId != None):
                    currentTime = datetime.datetime.now()
                    exc.inputTimestampIntoExcel(id, employeeId, currentTime)
                else:
                    exc.inputTokenIdToExce(id)
            #os.system('clear')
            tmp = abs(time.time() - lastReadedTimeStamp)
            print("LastReadedId: %s recieved %.2f seconds ago" % (hex(lastReadedId),tmp))
            if (tmp > 2):
                lastReadedId = 0

        #to delete   
        if(i == 3):
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


        
        
        
    



