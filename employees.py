#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Dec 28 20:45:07 2023

@author: wiktor
"""
##########################################################
# Imports
##########################################################
from excel import *

##########################################################
# Define
##########################################################
WORKING_STATUS_ERROR = 1
TRASH_ERROR = 2 
CHANGE_TO_Z_ERROR = 3
CHANGE_TO_T_OR_N_ERROR = 4
AMMOUNT_OF_EMPLOYEES_ERROR = 5

##########################################################
# Class
##########################################################
class employeesC:
    # --------------------------------------------------------------------
    # Class defines
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

    # --------------------------------------------------------------------
    # Functions
    def updateEmployeesStatus(self, xlsx):

        xlsx.checkFileAndInitialize()
        if(xlsx.isFileInitializeCorrectly):

            ammountOfEmploteesValue = \
                xlsx.workbook[DATA].cell(row = FIRST_DATA_ROW, column = AMMOUNT_OF_EMPLOYEES_COLUMN).value
            
            if (ammountOfEmploteesValue != None):
                
                #Settings function variables
                isCellFilled = True
                cellRow = FIRST_DATA_ROW
                cellsReaded = 0
                employeesToAdd = 0
                EmployeeStatusError = 0
                errorList = []

                #looping until found empty row
                while(isCellFilled):

                    # Reading mandatory data
                    tokenIdValue = xlsx.workbook[DATA].cell(row = cellRow, column = TOKEN_ID_COLUMN).value
                    surnameValue = xlsx.workbook[DATA].cell(row = cellRow, column = SURNAME_COLUMN).value
                    rateValue = xlsx.workbook[DATA].cell(row = cellRow, column = RATE_COLUMN).value
                    workingStatusValue = \
                            xlsx.workbook[DATA].cell(row = cellRow, column = WORKING_STATUS_COLUMN).value
                    
                    # New employee to add
                    if((workingStatusValue == 'Z') and (surnameValue != None)):
                        self.employee.append(self.employeeDataC(tokenIdValue, surnameValue, rateValue, workingStatusValue))
 
                        # Setting proper color after fixing error
                        self.restoreColor(xlsx, cellRow)

                        employeesToAdd += 1
                        cellsReaded += 1
                        cellRow += 1

                    elif ((tokenIdValue != None) and 
                          (surnameValue != None) and 
                          (rateValue !=  None) and
                          (workingStatusValue != None)):

                        # hired employees
                        if ((workingStatusValue == 'T') or (workingStatusValue == 'N')):
                            
                            self.employee.append(self.employeeDataC(tokenIdValue, \
                                                                    surnameValue, \
                                                                    rateValue, \
                                                                    workingStatusValue))
                            # Setting proper color after fixing error
                            self.restoreColor(xlsx, cellRow)

                            cellsReaded += 1
                            cellRow += 1

                        # Unknow employee status
                        else:
                            errorList.append([WORKING_STATUS_ERROR, cellRow, workingStatusValue])
                            EmployeeStatusError += 1
                            cellsReaded += 1
                            cellRow += 1
                            
                    elif((tokenIdValue != None) or 
                         (surnameValue != None) or 
                         (rateValue !=  None) or
                         (workingStatusValue != None)):
                        
                        errorList.append([TRASH_ERROR, cellRow, workingStatusValue])
                        EmployeeStatusError += 1
                        cellsReaded += 1
                        cellRow += 1
                
                    # No filled cells, end of loop
                    else:
                        isCellFilled = False
                
                ammountOfWorkingEmployee = 0
                # Checking ammount of employees
                if (EmployeeStatusError == 0):

                    ammountOfWorkingEmployee = (ammountOfEmploteesValue + employeesToAdd)

                    if (((cellsReaded) == ammountOfWorkingEmployee) and \
                        (len(self.employee) == ammountOfWorkingEmployee)):

                        #Sucsess 
                        xlsx.paintRows(DATA, FIRST_DATA_ROW, AMMOUNT_OF_EMPLOYEES_COLUMN, WHITE_PATTERN)
                        self.ammountOfEmplotees = ammountOfWorkingEmployee
                        self. restoreErrors(xlsx, cellsReaded)
                        self.isClassFilledCorectly = True
                        
                    else:
                        self.searchForErrors(xlsx, cellsReaded, ammountOfWorkingEmployee, errorList)
                else:
                    self.printErrors(xlsx, cellsReaded, ammountOfWorkingEmployee, errorList)
            else:
                logger.error("Ammount of employees cell is None")

        return self.isClassFilledCorectly


    def printErrors(self, xlsx, cellsReaded, ammountOfWorkingEmployee, errorList):
        for i in range(len(errorList)):
            rowToPasteErrorText = FIRST_DATA_ROW + cellsReaded + 1 + i

            # 1. WORKING_STATUS_ERROR
            if (errorList[i][0] == WORKING_STATUS_ERROR):

                xlsx.paintRows(DATA, errorList[i][1], TOKEN_ID_COLUMN, RED_PATTERN, \
                WORKING_STATUS_COLUMN)
                
                infoText = ("Błąd w rzędzie: %d, nieznany status pracownika" %errorList[i][1])
                xlsx.workbook[DATA].cell(row = rowToPasteErrorText, \
                    column = SURNAME_COLUMN).value = infoText
                
                logger.error("Unknown employee status %d" %errorList[i][1])
            # 2. TRASH_ERROR
            elif (errorList[i][0] == TRASH_ERROR):
                
                xlsx.paintRows(DATA, errorList[i][1], TOKEN_ID_COLUMN, RED_PATTERN, \
                                WORKING_STATUS_COLUMN)
                
                infoText = ("Błąd w rzędzie: %d, wykryto śmieciowe dane" %errorList[i][1])
                xlsx.workbook[DATA].cell(row = rowToPasteErrorText, \
                    column = SURNAME_COLUMN).value = infoText
                               
                logger.error("Trash data in row: %d" %errorList[i][1])

            # 3. CHANGE_TO_Z_ERROR
            elif (errorList[i][0] == CHANGE_TO_Z_ERROR):

                xlsx.paintRows(DATA, errorList[i][1], TOKEN_ID_COLUMN, RED_PATTERN, \
                               WORKING_STATUS_COLUMN)
                
                infoText = ("Błąd w rzędzie: %d, błędny satatus pracownika, zmień status na 'Z'" %errorList[i][1])
                xlsx.workbook[DATA].cell(row = rowToPasteErrorText, \
                    column = SURNAME_COLUMN).value = infoText
                               
                logger.error("Wrong employee status: %s in row: %d, propper status = 'Z'" 
                             %(errorList[i][2], errorList[i][1]))
            
            # 4. CHANGE_TO_T_OR_N_ERROR
            elif (errorList[i][0] == CHANGE_TO_T_OR_N_ERROR):

                xlsx.paintRows(DATA, errorList[i][1], TOKEN_ID_COLUMN, RED_PATTERN, \
                               WORKING_STATUS_COLUMN)

                infoText = ("Błąd w rzędzie: %d, błędny satatus pracownika, zmień status na 'T' lub 'N'" %errorList[i][1])
                xlsx.workbook[DATA].cell(row = rowToPasteErrorText, \
                    column = SURNAME_COLUMN).value = infoText
                               
                logger.error("Wrong employee status: %s in row: %d, propper status = 'Z'" 
                             %(errorList[i][2], errorList[i][1]))
            
            # 5. AMMOUNT_OF_EMPLOYEES_ERROR
            elif (errorList[i][0] == AMMOUNT_OF_EMPLOYEES_ERROR):
                
                xlsx.paintRows(DATA, FIRST_DATA_ROW, AMMOUNT_OF_EMPLOYEES_COLUMN, RED_PATTERN)

                infoText = ("Błędna ilość pracowników")
                xlsx.workbook[DATA].cell(row = rowToPasteErrorText, \
                    column = SURNAME_COLUMN).value = infoText
                               
                logger.error("Wrong employee status: %s in row: %d, propper status = 'Z'" 
                             %(errorList[i][2], errorList[i][1]))


    def searchForErrors(self, xlsx, cellsReaded, ammountOfWorkingEmployee, errorList):
        # Error statis information
        logger.error("Readed rows and ammount of employees is't the same. "
                                    "cellsReaded: " + str(cellsReaded) +
                                    ", ammountOfWorkingEmployee: " + str(ammountOfWorkingEmployee) +
                                    ", len(self.employee): " + str(len(self.employee)))

        estimatedPotentialProblems = abs(cellsReaded - ammountOfWorkingEmployee)
        logger.error("Employee status update: Error, estimated potential problems: %d, searching for potential problem..." %(estimatedPotentialProblems))
        potentialProblems = 0

        for i in range(cellsReaded):
            rowToEdit = i + FIRST_DATA_ROW
            hoursValue = xlsx.workbook[DATA].cell(row = rowToEdit, \
                column = HOURS_COLUMN).value
            salaryValue = xlsx.workbook[DATA].cell(row = rowToEdit, \
                column = SALARY_COLUMN).value
            workingStatusValue = xlsx.workbook[DATA].cell(row = rowToEdit, \
                column = WORKING_STATUS_COLUMN).value
            
            if (workingStatusValue == 'T' or workingStatusValue == 'N'):
                if (hoursValue == None or salaryValue == None):
                    errorList.append([CHANGE_TO_Z_ERROR, rowToEdit, workingStatusValue])
                    potentialProblems += 1

            elif (workingStatusValue == 'Z'):
                if (hoursValue != None or salaryValue != None):
                    errorList.append([CHANGE_TO_T_OR_N_ERROR, rowToEdit, workingStatusValue])
                    potentialProblems += 1

        if (potentialProblems == 0):
            errorList.append([AMMOUNT_OF_EMPLOYEES_ERROR, rowToEdit, workingStatusValue])

        self.printErrors(xlsx, cellsReaded, ammountOfWorkingEmployee, errorList)


    def restoreColor(self, xlsx, cellRow):
        rowColor = xlsx.workbook[DATA].cell(row=cellRow, column=TOKEN_ID_COLUMN).\
                    fill.bgColor.index
        rowColor2 = xlsx.workbook[DATA].cell(row=cellRow, column=RATE_COLUMN).\
                    fill.bgColor.index
        if (((cellRow % 2) == 0) and ((rowColor != 'FFC0C0C0') or (rowColor2 != 'FFC0C0C0'))):
            xlsx.paintRows(DATA, \
                            cellRow, \
                            TOKEN_ID_COLUMN, \
                            GREY_PATTERN, \
                            WORKING_STATUS_COLUMN)
            
        if (((cellRow % 2) == 1) and ((rowColor != '00000000') or (rowColor2 != '00000000'))):
            xlsx.paintRows(DATA, \
                            cellRow, \
                            TOKEN_ID_COLUMN, \
                            WHITE_PATTERN, \
                            WORKING_STATUS_COLUMN)

    def restoreErrors(self, xlsx, cellsReaded):

        keepSearching = True
        cellRow = FIRST_DATA_ROW + cellsReaded

        rowColor = xlsx.workbook[DATA].cell(row=cellRow, column=TOKEN_ID_COLUMN).\
                    fill.bgColor.index
        if (rowColor != '00000000'):
            for i in range (WORKING_STATUS_COLUMN):
                xlsx.workbook[DATA].cell(row=cellRow, column=i+1).style = 'Normal'
        
        emptyRows = 0   
        while (keepSearching):

            cellRow += 1
            emptyRows += 1
            rowColor = xlsx.workbook[DATA].cell(row=cellRow, column=TOKEN_ID_COLUMN).\
                        fill.bgColor.index
            rowValue = xlsx.workbook[DATA].cell(row=cellRow, column=SURNAME_COLUMN).value
            if (rowColor != '00000000'):
                for i in range (WORKING_STATUS_COLUMN):
                    xlsx.workbook[DATA].cell(row=cellRow, column=i+1).style = 'Normal'
                emptyRows = 0

            if (rowValue != None):
                xlsx.workbook[DATA].cell(row=cellRow, column=SURNAME_COLUMN).value = None

                emptyRows = 0

            if (emptyRows == 2):
                keepSearching = False


    def deleteLayedEmployees(self):
        i = 0
        while (i < self.ammountOfEmplotees):
            if(self.employee[i].workingStatus == 'N'):
                self.employee.pop(i)
                self.ammountOfEmplotees -= 1
            else:
                i += 1


    def addNewEmployees(self, xlsx, now):
        ammountOfNewEmployees = 0
        for i in range(self.ammountOfEmplotees):
            editedRow = FIRST_DATA_ROW + i
            if(self.employee[i].workingStatus == 'Z'):
                ammountOfNewEmployees += 1
                filledCells = 0
                tokenIdValue = xlsx.workbook[DATA].cell(row = editedRow, column = TOKEN_ID_COLUMN).value
                rateValue = xlsx.workbook[DATA].cell(row = editedRow, column = RATE_COLUMN).value

                if (tokenIdValue != None):
                    filledCells += 1
                else:
                    xlsx.paintRows(DATA, editedRow, TOKEN_ID_COLUMN, ORANGE_PATTERN)
                
                if (rateValue != None):
                    filledCells += 1
                else:
                    xlsx.paintRows(DATA, editedRow, RATE_COLUMN, ORANGE_PATTERN)

                if (filledCells == 2):
                    xlsx.addNewEmployeeCells(i, self.employee[i], now)
                    xlsx.updateAmmountOfEmployees(self.ammountOfEmplotees)
                    ammountOfNewEmployees -= 1

        return ammountOfNewEmployees
                

    def checkReadedId(self, id):
        isIdAssignedToEmployee = None
        for i in range(0, self.ammountOfEmplotees):
            if (id == self.employee[i].tokenId):
                isIdAssignedToEmployee =  i
                break
        return isIdAssignedToEmployee
                