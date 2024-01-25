#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Dec 28 20:45:07 2023

@author: wiktor
"""


import logging
from excel import *
FORMAT = '%(asctime)s %(levelname)s at line %(lineno)s %(name)s %(funcName)s : %(message)s'
logging.basicConfig(filename='myapp.log', level=logging.DEBUG, format = FORMAT, force=True)
del FORMAT
logger=logging.getLogger(__name__)
employeeFile = 'employee.yaml'


class employeesC:
    # -> Class declaration >---------------------------------------------------
    def __init__(self):
        self.employee = []
        self.ammountOfEmplotees = 0
        self.isClassFilledCorectly = False

    class employeeDataC:
        def __init__(self, tokenid, name, rate, workingStatus):
            self.name = name
            self.rate = rate
            self.tokenId = tokenid
            self.workingStatus = workingStatus

    def updateEmployeesStatus(self, xlsx):
        xlsx.checkFileAndInitialize()
        if(xlsx.isFileInitializeCorrectly): #TO DO add error exception
            ammountOfEmploteesValue = \
                xlsx.workbook[DATA].cell(row = FIRST_DATA_ROW, column = AMMOUNT_OF_EMPLOYEES_COLUMN).value
            if (ammountOfEmploteesValue != None):

                isCellFilled = True
                cellRow = FIRST_DATA_ROW
                cellsReaded = 0
                employeesToAdd = 0
                notValidCells = 0

                while(isCellFilled):
                    # Reading mandatory data
                    tokenIdValue = xlsx.workbook[DATA].cell(row = cellRow, column = TOKEN_ID_COLUMN).value
                    surnameValue = xlsx.workbook[DATA].cell(row = cellRow, column = SURNAME_COLUMN).value
                    rateValue = xlsx.workbook[DATA].cell(row = cellRow, column = RATE_COLUMN).value
                    workingStatusValue = \
                            xlsx.workbook[DATA].cell(row = cellRow, column = WORKING_STATUS_COLUMN).value
                    
                    if ((tokenIdValue != None) and 
                        (surnameValue != None) and 
                        (rateValue !=  None) and
                        (workingStatusValue != None)):
                        tmp = xlsx.workbook[DATA].cell(row=cellRow, column=(TOKEN_ID_COLUMN)).fill.bgColor.index
                        if ((workingStatusValue == 'T') or (workingStatusValue == 'N')):
                            self.employee.append(self.employeeDataC(tokenIdValue, surnameValue, rateValue, workingStatusValue))
                            if (xlsx.workbook[DATA].cell(row=cellRow, column=(TOKEN_ID_COLUMN)).fill.bgColor.index == 'FF993300'):
                                xlsx.paintRows(DATA, cellRow, TOKEN_ID_COLUMN, WHITE_PATTERN, WORKING_STATUS_COLUMN)
                        elif (workingStatusValue == 'Z'):
                           self.employee.append(self.employeeDataC(tokenIdValue, surnameValue, rateValue, workingStatusValue))
                           employeesToAdd += 1
                        else:
                            notValidCells += 1
                            xlsx.paintRows(DATA, cellRow, TOKEN_ID_COLUMN, RED_PATTERN, WORKING_STATUS_COLUMN)
                            logger.error("Unknown employee working status. Cell: " + 
                                         str(cellRow) + " workingStatusValue: " + 
                                         workingStatusValue)
                            
                        cellsReaded += 1
                        cellRow += 1
                    else:
                        isCellFilled = False
                if (notValidCells == 0):

                    ammountOfWorkingEmployee = (ammountOfEmploteesValue + employeesToAdd)

                    if (((cellsReaded) == ammountOfWorkingEmployee) and \
                        (len(self.employee) == ammountOfWorkingEmployee)):

                        #Sucsess 
                        self.ammountOfEmplotees = ammountOfWorkingEmployee
                        print("Employee status update: suscess")
                        self.isClassFilledCorectly = True
                    else:
                        logger.error("Readed rows and ammount of employees is't the same. "
                                    "cellsReaded: " + str(cellsReaded) +
                                    " ammountOfWorkingEmployee: " + str(ammountOfWorkingEmployee) +
                                    " len(self.employee): " + str(len(self.employee)))

                        estimatedPotentialProblems = abs(cellsReaded - ammountOfWorkingEmployee)
                        print("Employee status update: Error, estimated potential problems: %d, searching for potential problem..." %(estimatedPotentialProblems))
                        potentialProblems = 0

                        for i in range(cellsReaded):
                            rowToEdit = i + FIRST_DATA_ROW
                            hoursValue = xlsx.workbook[DATA].cell(row = rowToEdit, \
                                column = HOURS_COLUMN).value
                            salaryValue = xlsx.workbook[DATA].cell(row = rowToEdit, \
                                column = SALARY_COLUMN).value
                            workingStatusValue = xlsx.workbook[DATA].cell(row = rowToEdit, \
                                column = WORKING_STATUS_COLUMN).value
                            
                            if (hoursValue == None or salaryValue == None):
                                potentialProblems += 1

                                xlsx.paintRows(DATA, rowToEdit, TOKEN_ID_COLUMN, RED_PATTERN, \
                                               WORKING_STATUS_COLUMN)
                                
                                print("Potential problems find: %d of %d " \
                                      %(potentialProblems, estimatedPotentialProblems))
                                logger.error("Probably problem find in row: " + str(rowToEdit) +
                                             " hoursValue: " + str(hoursValue) +
                                             " salaryValue: " + str(salaryValue) +
                                             " workingStatusValue: " + str(workingStatusValue))
                                
                                if(workingStatusValue != "Z"):
                                    rowToPasteError = TITLE_ROW + \
                                                        cellsReaded + \
                                                        potentialProblems + \
                                                        2
                                    
                                    valueToInput = "BŁĄD W RZĘDZIE: " + str(rowToEdit) + ", PRAWDOPODOBNIE BŁĘDNY STATUS PRACOWNIKA, ZMIEŃ STATUS NA 'Z'."
                                    xlsx.workbook[DATA].cell(row = rowToPasteError, \
                                        column = SURNAME_COLUMN).value = valueToInput
                            
                            elif(workingStatusValue == "Z"):
                                potentialProblems += 1
                                xlsx.paintRows(DATA, rowToEdit, TOKEN_ID_COLUMN, RED_PATTERN, \
                                               WORKING_STATUS_COLUMN)
                                print("Potential problems find: %d of %d " \
                                      %(potentialProblems, estimatedPotentialProblems))
                                
                                rowToPasteError = TITLE_ROW + \
                                                    cellsReaded + \
                                                    potentialProblems + \
                                                    2
                                    
                                valueToInput = "BŁĄD W RZĘDZIE: " + str(rowToEdit) + ", PRAWDOPODOBNIE BŁĘDNY STATUS PRACOWNIKA, ZMIEŃ STATUS NA 'T' LUB 'N'."
                                xlsx.workbook[DATA].cell(row = rowToPasteError, \
                                    column = SURNAME_COLUMN).value = valueToInput
                                    
            else:
                logger.error("Ammount of employees cell is None")
        return self.isClassFilledCorectly
            
     
    def deleteLayedEmployees(self):
        i = 0
        while (i < self.ammountOfEmplotees):
            if(self.employee[i].workingStatus == 'N'):
                self.employee.pop(i)
                self.ammountOfEmplotees -= 1
            else:
                i += 1

    def addNewEmployees(self, xlsx, now):
        for i in range(self.ammountOfEmplotees):
            if(self.employee[i].workingStatus == 'Z'):
                xlsx.addNewEmployeeCells(i, self.employee[i], now)
                xlsx.updateAmmountOfEmployees(self.ammountOfEmplotees)
                

    