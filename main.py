# -*- coding: utf-8 -*-

##########################################################
# Imports
##########################################################
from excel import *
from  employees import *
import datetime
import os
import gpio
from mfrc522 import SimpleMFRC522
reader = SimpleMFRC522()

##########################################################
# Defines
##########################################################
STARTING_WORK_TIME = 7
COOLDOWN_TIME_IN_SEC = 2

##########################################################
# Functions
##########################################################
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


def checkMonthAndFile(employees, exc, now, currentMonth, isSafeMode):
    if (now.month != currentMonth):
        gpio.excelProcessed()
        isSafeMode = True
        employeesAfterUpdate = employeesC()
        lastExcelFile = exc.filename
        if(employeesAfterUpdate.updateEmployeesStatus(exc)):
            if(employeesAfterUpdate.addNewEmployees(exc, now) == 0):
                employeesAfterUpdate.deleteLayedEmployees()
                employees = employeesAfterUpdate 
                currentExcelFile = str(now.month) + "_" + str(now.year) + ".xlsx"
                exc = changeFile(employees, exc, now, currentExcelFile)
                try:
                    os.replace(DIRECTORY + lastExcelFile, DIRECTORY + ARCHIVE + lastExcelFile)
                except Exception as e: 
                    pass
                currentMonth = now.month
                isSafeMode = False
            else:
                logger.error("Theres still new employees in sheet.")
    return employees, exc, currentMonth, isSafeMode


def updateEmployees(employees, exc, now):
    gpio.excelProcessed()
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

##########################################################
# Main
##########################################################
if __name__=="__main__": 

    ########################################################################
    #TEST
        #currentExcelFile = '2_2024.xlsx'
    NOWYEAR = 2024
    NOWMONTH = 2
    NOWDAY = 1
    NOWHOUR = 1
    NOWMINUTES = 0
    NOWSECONDS = 0
    now = datetime.datetime(NOWYEAR, NOWMONTH, NOWDAY, NOWHOUR, NOWMINUTES, NOWSECONDS)
    ####################################################################
    
    # starting
    logger.error("Program start running...")
    gpio.gpioInitialize()
    isSafeMode = False

    # Time settings
    #now = datetime.datetime.now() TEST
    currentMonth = now.month
    lastSecond = 0

    # Generating objects
    employees = employeesC()
    file = str(now.month) + "_" + str(now.year) + ".xlsx"
    exc = excelSheetC(file)

    # Generating template if file dosn't exist
    exc.checkFileAndInitialize()
    if(not(exc.isFileInitializeCorrectly)):
        isSafeMode = True

        excTmp = excelSheetC(str(now.month) + "_" + str(now.year) + "(Template)" + ".xlsx")
        excTmp.create_workbook()
        excTmp.generateExcelTemplate(employees, now)
        excTmp.saveSheet()
        del(excTmp)

    #setting steering parametes
    lastReadedId = 0
    lastReadedTimeStamp = 0
    coolDown = 0
    signalLed = True

    # main loop
    while(True):
        # setting current time
        #now = datetime.datetime.now()
        # checking month and upating file if theres any changes
        if(os.path.isfile(exc.file)):
            employees, exc, currentMonth, isSafeMode = \
                checkMonthAndFile(employees, exc, now, currentMonth, isSafeMode)
            modificationTimeSample = time.ctime(os.path.getmtime(exc.file))
            if (exc.lastModificationTime != modificationTimeSample):
                employees, exc, isSafeMode = updateEmployees(employees, exc, now)

        # reading token
        readedTokenId = reader.read_no_block()[0]
        # processing token
        if ((readedTokenId != None) and (readedTokenId != lastReadedId)):
            gpio.readingTokenLed()
            
            # setting cooldown
            lastReadedTimeStamp = time.time()
            lastReadedId = readedTokenId

            # changing time to 7 is its earlier
            if (now.hour != (STARTING_WORK_TIME - 1)):
                timeToPasteInExcel = now
            else:
                timeToPasteInExcel = \
                    now.replace(hour = STARTING_WORK_TIME, minute=0, second= 0, microsecond= 0) 

            # putting token in excel 
            if (isSafeMode):
                logger.error("READED TOKENID: " + str(readedTokenId) + " at time: " + str(now))
                gpio.blinkSuccess()
            else:
                employeeId = employees.checkReadedId(readedTokenId)
                if (employeeId != None):
                    exc.inputTimestampIntoExcel(readedTokenId, employeeId, timeToPasteInExcel)
                    gpio.blinkSuccess()
                else:
                    exc.inputTokenIdToExce(readedTokenId)
                    gpio.blinkFailiture()
        # taking of cooldown
        coolDown = abs(time.time() - lastReadedTimeStamp)
        if (coolDown > COOLDOWN_TIME_IN_SEC):
            lastReadedId = 0
        
        # blink
        

        now2 = datetime.datetime.now()
        now = datetime.datetime(NOWYEAR, NOWMONTH, NOWDAY, NOWHOUR, NOWMINUTES, now2.second)

        #signalLed, lastSecond = gpio.blink(now, lastSecond, isSafeMode, signalLed)
        #TEST 
        if(now2.second != lastSecond):
            lastSecond = now.second
            NOWDAY +=1
            if (NOWMINUTES == 60):
                NOWMINUTES = 0
                NOWHOUR += 1
            if (NOWHOUR >= 24):
                NOWHOUR = 1
                NOWDAY += 1
            if(NOWDAY > 27):
                NOWDAY = 1
                NOWMONTH += 1
                    
            if(NOWMONTH > 12):
                NOWMONTH = 1
                NOWYEAR +=1

            print(str(now))

        