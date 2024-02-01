# -*- coding: utf-8 -*-

from excel import *
from  employees import *
import datetime
import os
import time
import RPi.GPIO as GPIO
GPIO.setmode(GPIO.BCM)


BUZZER = 12
YELLOW_LED = 20
GREEN_LED = 21

GPIO.setup(BUZZER, GPIO.OUT)
GPIO.setup(YELLOW_LED,GPIO.OUT)
GPIO.setup(GREEN_LED,GPIO.OUT)

GPIO.output(BUZZER,False)
GPIO.output(YELLOW_LED,False)
GPIO.output(GREEN_LED,False)

from mfrc522 import SimpleMFRC522
reader = SimpleMFRC522()
COOLDOWN_TIME_IN_SEC = 2

def blinkSuccess(ledPort):
    GPIO.output(YELLOW_LED, False)
    GPIO.output(GREEN_LED, False)
    
    GPIO.output(ledPort, True)
    GPIO.output(BUZZER, True)
    time.sleep(0.05)

    GPIO.output(ledPort, False)
    GPIO.output(BUZZER, False)
    time.sleep(0.05)

    GPIO.output(ledPort, True)
    GPIO.output(BUZZER, True)
    time.sleep(0.05)
    
    GPIO.output(ledPort, False)
    GPIO.output(BUZZER, False)

def blinkFailiture(ledPort):

    GPIO.output(YELLOW_LED, False)
    GPIO.output(GREEN_LED, False)
    
    GPIO.output(ledPort, True)
    GPIO.output(BUZZER, True)
    time.sleep(0.05)

    GPIO.output(ledPort, False)
    time.sleep(0.05)

    GPIO.output(ledPort, True)
    time.sleep(0.05)
    
    GPIO.output(ledPort, False)
    time.sleep(0.05)

    GPIO.output(ledPort, True)
    time.sleep(0.05)
    
    GPIO.output(ledPort, False)
    time.sleep(0.05)
    
    GPIO.output(BUZZER, False)

def readingTokenLed():
    GPIO.output(YELLOW_LED, True)
    GPIO.output(GREEN_LED, False)
    
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
        print("Changing month")
        isSafeMode = True
        employeesAfterUpdate = employeesC()
        if(employeesAfterUpdate.updateEmployeesStatus(exc)):
            if(employeesAfterUpdate.addNewEmployees(exc, now) == 0):
                employeesAfterUpdate.deleteLayedEmployees()
                employees = employeesAfterUpdate 

                currentExcelFile = str(now.month) + "_" + str(now.year) + ".xlsx"
                exc = changeFile(employees, exc, now, currentExcelFile)
                currentMonth = now.month
                isSafeMode = False
            else:
                print("Theres still new employees in sheet.")
                logger.error("Theres still new employees in sheet.")
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
    logger.error("Program start running...")   
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

    employees = employeesC()
    exc = excelSheetC(str(now.month) + "_" + str(now.year) + ".xlsx")
    exc.checkFileAndInitialize()
    if(not(exc.isFileInitializeCorrectly)):
        isSafeMode = True

        excTmp = excelSheetC(str(now.month) + "_" + str(now.year) + "(Template)" + ".xlsx")
        excTmp.create_workbook()
        excTmp.generateExcelTemplate(employees, now)
        excTmp.saveSheet()
        del(excTmp)


    i = 0
    #updateEmployees(exc, now)
    lastReadedId = 0
    lastReadedTimeStamp = 0
    coolDown = 0
    signalLed = True
    signalLedPort = GREEN_LED
    unusedSignalLedPort = YELLOW_LED
    while(True):

        if(os.path.isfile(exc.file)):
            employees, exc, currentMonth, isSafeMode = \
                checkMonthAndFile(employees, exc, now, currentMonth, isSafeMode)
            modificationTimeSample = time.ctime(os.path.getmtime(exc.file))
            if (exc.lastModificationTime != modificationTimeSample):
                employees, exc, isSafeMode = updateEmployees(employees, exc, now)

        if (isSafeMode):
            signalLedPort = YELLOW_LED
            unusedSignalLedPort = GREEN_LED
        else:
            signalLedPort = GREEN_LED
            unusedSignalLedPort = YELLOW_LED

        readedTokenId = reader.read_no_block()[0]
        if ((readedTokenId != None) and (readedTokenId != lastReadedId)):
            #Led steering
            readingTokenLed()
            
            lastReadedTimeStamp = time.time()
            lastReadedId = readedTokenId

            currentTime = now #datetime.datetime.now()
            if (isSafeMode):
                logger.error("READED TOKENID: " + str(readedTokenId) + " at time: " + str(currentTime))
                blinkSuccess(GREEN_LED)
            else:
                employeeId = employees.checkReadedId(readedTokenId)
                if (employeeId != None):
                    exc.inputTimestampIntoExcel(readedTokenId, employeeId, currentTime)
                    blinkSuccess(GREEN_LED)
                else:
                    exc.inputTokenIdToExce(readedTokenId)
                    blinkFailiture(YELLOW_LED)
            
            print("LastReadedId: %d" %lastReadedId)
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
            tmptext = ''
            if(isSafeMode):
                tmptext = ' SafeMode enabled'
            print(str(now) + tmptext)
            lastMinute = now2.second
            NOWMINUTES += 15
            signalLed = not(signalLed)
            GPIO.output(signalLedPort, signalLed)
            GPIO.output(unusedSignalLedPort, False)
