## Bibliothéque du fonctions

from ast import Continue, pattern
from itertools import count
##base du donneés json
from json import load
from multiprocessing.sharedctypes import Value
from pipes import Template
from tkinter.tix import CELL
from types import CellType
## excel
from openpyxl import Workbook, load_workbook, cell
from pyparsing import empty
## base du donneés
from tinydb import TinyDB, Query, where
import json
import re
from EmploiApp import Ui_MainWindow


##creé base du donneés
def DBcreate(): 
 db = TinyDB('db3333.json')
 return db
"""
cmd2='python testbutton.py'
p=subprocess.Popen(cmd2, shell=True)
out, err =p.communicate()
print(err)
print(out)
"""
##excel file
def LoadXLFile(FileName):
    WB = load_workbook(FileName)
    WS = WB.active
    return WS


##Main function cacule max lignes
def searchLastRow(sheet, base):
    for row in range(base, sheet.max_row):
        if sheet.cell(row=row, column=1).value is None: continue
        if next((filter(lambda x: x is not None, [
                sheet.cell(row=i, column=1).value
                for i in range(row + 1, row + 101)
        ])), None) == None:
            return row


##List debut creneaux et fin creneau

## finds all the starts and ends

def parcourirCreneau(WS):
    CoordinatesList = []
    base = 4
    lastrow = searchLastRow(WS, 4)
    row = 4
    table_creneaux = []

    while row <= lastrow:
        creneau = {"debut": WS.cell(row=row, column=1).value, "indice-debut": row}
        row += 1

        while row <= lastrow:
            if WS.cell(row=row, column=1).value is not None:
                break
            row += 1

        end = WS.cell(row=row, column=1).value
        creneau["fin"] = WS.cell(row=row, column=1).value
        creneau["indice-fin"] = row
        table_creneaux.append(creneau)
        CoordinatesList.append(row)
    ##print(creneau)
        row += 1

#list of the sessions
    print(CoordinatesList)

    seances1 = []
    seances2 = []
    seances3 = []
    seances4 = []
    seances5 = []
    seances6 = []
    seances7 = []
    seances8 = []
    seances9 = []
    seances10 = []
    seances11 = []
    seances12 = []
    seances13 = []
    seances14 = []
    seances15 = []
    seances16 = []
    seances17 = []
    seances18 = []
    seances19 = []
    seances20 = []
    seances21 = []
    seances22 = []
    seances23 = []
    seances24 = []
    seances25 = []

##this will become a button in the interface

##Day1
## from 1 - 5 picks the creneaus of the day

##print((CoordinatesList[0]-sizeofcreneau)+counter)
##print(sizeofcreneau)
##print(WS.cell(row=(CoordinatesList[0]-sizeofcreneau)+counter,column=2).value)

    counter = 1
    countercours = 1
    sizeofcreneau = (CoordinatesList[1] + 1) - (CoordinatesList[0] + 1)
## DAY 1
## CRENEAU 1
    if WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
           column=2).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
                    column=2).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[0] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[0],
                            min_col=2,
                            max_col=6)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[0] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[0],
                             min_col=2,
                             max_col=2)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances1.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
            ##print(counter)
                if countercours == 1:
                    seances1.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances1.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances1.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances1.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances1.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances1.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
                           column=2).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) +
                                    counter,
                                    column=2).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) +
                                    counter,
                                    column=2).value)) is None:
                        countercours = 1

    print("Creneau 1 worked")

#CRENAU 2
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
           column=2).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
                    column=2).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau2 = WS.iter_rows(min_row=(CoordinatesList[1] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[1],
                            min_col=2,
                            max_col=6)
        CreneauY2 = WS.iter_rows(min_row=(CoordinatesList[1] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[1],
                             min_col=2,
                             max_col=2)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau2:
                    counter += 1
                    seances2.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY2:
                print(counter)
                if countercours == 1:
                    seances2.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances2.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances2.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances2.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances2.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances2.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
                           column=2).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) +
                                    counter,
                                    column=2).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) +
                                    counter,
                                    column=2).value)) is None:
                        countercours = 1

    print("Creneau 2 worked")
##CRENNEAU 3
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
           column=2).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
                    column=2).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[2] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[2],
                            min_col=2,
                            max_col=6)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[2] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[2],
                             min_col=2,
                             max_col=2)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances3.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances3.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances3.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances3.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances3.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances3.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances3.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
                           column=2).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) +
                                    counter,
                                    column=2).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) +
                                    counter,
                                    column=2).value)) is None:
                        countercours = 1
    print("Creneau 3 worked")
##CRENEAU 4
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
           column=2).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
                    column=2).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[3] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[3],
                            min_col=2,
                            max_col=6)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[3] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[3],
                             min_col=2,
                             max_col=2)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances4.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances4.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances4.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances4.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances4.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances4.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances4.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
                           column=2).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) +
                                    counter,
                                    column=2).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) +
                                    counter,
                                    column=2).value)) is None:
                        countercours = 1

    print("Creneau 4 worked")

#CRENEAU 5
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
           column=2).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
                    column=2).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[4] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[4],
                            min_col=2,
                            max_col=6)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[4] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[4],
                             min_col=2,
                             max_col=2)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances5.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances5.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances5.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances5.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances5.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances5.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances5.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
                           column=2).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) +
                                    counter,
                                    column=2).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) +
                                    counter,
                                    column=2).value)) is None:
                        countercours = 1

    print("Creneau 5 worked")
## Day 2
##creneau 1
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
           column=7).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
                    column=7).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[0] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[0],
                            min_col=7,
                            max_col=11)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[0] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[0],
                             min_col=7,
                             max_col=7)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances6.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances6.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances6.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances6.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances6.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances6.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances6.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
                           column=7).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) +
                                    counter,
                                    column=7).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) +
                                    counter,
                                    column=7).value)) is None:
                        countercours = 1

    print("Creneau 1 day 2 worked")

#CRENAU 2
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
           column=7).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
                    column=7).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[1] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[1],
                            min_col=7,
                            max_col=11)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[1] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[1],
                             min_col=7,
                             max_col=7)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances7.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances7.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances7.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances7.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances7.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances7.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances7.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
                           column=7).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) +
                                    counter,
                                    column=7).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) +
                                    counter,
                                    column=7).value)) is None:
                        countercours = 1

    print("Creneau 2 worked")
##CRENNEAU 3
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
           column=7).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
                    column=7).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[2] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[2],
                            min_col=7,
                            max_col=11)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[2] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[2],
                             min_col=7,
                             max_col=7)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances8.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances8.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances8.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances8.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances8.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances8.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances8.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
                           column=7).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) +
                                    counter,
                                    column=7).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) +
                                    counter,
                                    column=7).value)) is None:
                        countercours = 1

    print("Creneau 3 worked")
##CRENEAU 4
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
           column=7).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
                    column=7).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[3] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[3],
                            min_col=7,
                            max_col=11)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[3] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[3],
                             min_col=7,
                             max_col=7)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances9.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances9.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances9.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances9.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances9.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances9.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances9.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
                           column=7).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) +
                                    counter,
                                    column=7).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) +
                                    counter,
                                    column=7).value)) is None:
                        countercours = 1

    print("Creneau 4 worked")

#CRENEAU 5
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
           column=7).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
                    column=7).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[4] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[4],
                            min_col=7,
                            max_col=11)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[4] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[4],
                             min_col=7,
                             max_col=7)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances10.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances10.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances10.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances10.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances10.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances10.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances10.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
                           column=7).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) +
                                    counter,
                                    column=7).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) +
                                    counter,
                                    column=7).value)) is None:
                        countercours = 1

    print("Creneau 5 worked")
## Day 3
##creneau 1
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
           column=12).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
                    column=12).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[0] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[0],
                            min_col=12,
                            max_col=16)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[0] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[0],
                             min_col=12,
                             max_col=12)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances11.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances11.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances11.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances11.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances11.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances11.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances11.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
                           column=12).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) +
                                    counter,
                                    column=12).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) +
                                    counter,
                                    column=12).value)) is None:
                        countercours = 1

#CRENAU 2
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
           column=12).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
                    column=12).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[1] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[1],
                            min_col=12,
                            max_col=16)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[1] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[1],
                             min_col=12,
                             max_col=12)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances12.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances12.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances12.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances12.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances12.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances12.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances12.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
                           column=12).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) +
                                    counter,
                                    column=12).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) +
                                    counter,
                                    column=12).value)) is None:
                        countercours = 1

##CRENNEAU 3
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
           column=12).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
                    column=12).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[2] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[2],
                            min_col=12,
                            max_col=16)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[2] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[2],
                             min_col=12,
                             max_col=12)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances13.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances13.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances13.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances13.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances13.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances13.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances13.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
                           column=12).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) +
                                    counter,
                                    column=12).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) +
                                    counter,
                                    column=12).value)) is None:
                        countercours = 1

##CRENEAU 4
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
           column=12).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
                    column=12).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[3] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[3],
                            min_col=12,
                            max_col=16)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[3] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[3],
                             min_col=12,
                             max_col=12)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances14.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances14.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances14.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances14.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances14.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances14.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances14.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
                           column=12).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) +
                                    counter,
                                    column=12).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) +
                                    counter,
                                    column=12).value)) is None:
                        countercours = 1

#CRENEAU 5
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
           column=12).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
                    column=12).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[4] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[4],
                            min_col=12,
                            max_col=16)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[4] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[4],
                             min_col=12,
                             max_col=12)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances15.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances15.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances15.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances15.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances15.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances15.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances15.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
                           column=12).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) +
                                    counter,
                                    column=12).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) +
                                    counter,
                                    column=12).value)) is None:
                        countercours = 1

## Day 4
##creneau 1
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
           column=17).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
                    column=17).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[0] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[0],
                            min_col=17,
                            max_col=21)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[0] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[0],
                             min_col=17,
                             max_col=17)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances16.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances16.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances16.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances16.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances16.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances16.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances16.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
                           column=17).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) +
                                    counter,
                                    column=17).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) +
                                    counter,
                                    column=17).value)) is None:
                        countercours = 1

#CRENAU 2
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
           column=17).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
                    column=17).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[1] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[1],
                            min_col=17,
                            max_col=21)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[1] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[1],
                             min_col=17,
                             max_col=17)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances17.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances17.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances17.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances17.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances17.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances17.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances17.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
                           column=17).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) +
                                    counter,
                                    column=17).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) +
                                    counter,
                                    column=17).value)) is None:
                        countercours = 1

##CRENNEAU 3
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
           column=17).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
                    column=17).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[2] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[2],
                            min_col=17,
                            max_col=21)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[2] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[2],
                             min_col=17,
                             max_col=17)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances18.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances18.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances18.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances18.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances18.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances18.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances18.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
                           column=17).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) +
                                    counter,
                                    column=17).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) +
                                    counter,
                                    column=17).value)) is None:
                        countercours = 1

##CRENEAU 4
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
           column=17).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
                    column=17).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[3] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[3],
                            min_col=17,
                            max_col=21)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[3] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[3],
                             min_col=17,
                             max_col=17)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances19.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances19.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances19.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances19.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances19.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances19.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances19.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
                           column=17).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) +
                                    counter,
                                    column=17).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) +
                                    counter,
                                    column=17).value)) is None:
                        countercours = 1

#CRENEAU 5
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
           column=17).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
                    column=17).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[4] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[4],
                            min_col=17,
                            max_col=21)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[4] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[4],
                             min_col=17,
                             max_col=17)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances20.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances20.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances20.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances20.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances20.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances20.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances20.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
                           column=17).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) +
                                    counter,
                                    column=17).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) +
                                    counter,
                                    column=17).value)) is None:
                        countercours = 1

## Day 5
##creneau 1
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
           column=22).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
                    column=22).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[0] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[0],
                            min_col=22,
                            max_col=26)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[0] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[0],
                             min_col=22,
                             max_col=22)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances21.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances21.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances21.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances21.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances21.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances21.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances21.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[0] - sizeofcreneau) + counter,
                           column=22).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) +
                                    counter,
                                    column=22).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[0] - sizeofcreneau) +
                                    counter,
                                    column=22).value)) is None:
                        countercours = 1

#CRENAU 2
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
           column=22).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
                    column=22).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[1] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[1],
                            min_col=22,
                            max_col=26)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[1] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[1],
                             min_col=22,
                             max_col=22)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances22.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances22.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances22.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances22.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances22.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances22.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances22.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[1] - sizeofcreneau) + counter,
                           column=22).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) +
                                    counter,
                                    column=22).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[1] - sizeofcreneau) +
                                    counter,
                                    column=22).value)) is None:
                        countercours = 1

##CRENNEAU 3
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
           column=22).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
                    column=22).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[2] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[2],
                            min_col=22,
                            max_col=26)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[2] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[2],
                             min_col=22,
                             max_col=22)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances23.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances23.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances23.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances23.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances23.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances23.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances23.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[2] - sizeofcreneau) + counter,
                           column=22).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) +
                                    counter,
                                    column=22).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[2] - sizeofcreneau) +
                                    counter,
                                    column=22).value)) is None:
                        countercours = 1

##CRENEAU 4
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
           column=22).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
                    column=22).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[3] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[3],
                            min_col=22,
                            max_col=26)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[3] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[3],
                             min_col=22,
                             max_col=22)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances24.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances24.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances24.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances24.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances24.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances24.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances24.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[3] - sizeofcreneau) + counter,
                           column=22).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) +
                                    counter,
                                    column=22).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[3] - sizeofcreneau) +
                                    counter,
                                    column=22).value)) is None:
                        countercours = 1

#CRENEAU 5
    counter = 1
    countercours = 1
    if WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
           column=22).value is not None:
        result = re.search(
        "^G[1-9]$",
        str(
            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
                    column=22).value))
    else:
        result = None

    while counter < sizeofcreneau:
        Creneau1 = WS.iter_rows(min_row=(CoordinatesList[4] - sizeofcreneau) +
                            counter,
                            max_row=CoordinatesList[4],
                            min_col=22,
                            max_col=26)
        CreneauY1 = WS.iter_rows(min_row=(CoordinatesList[4] - sizeofcreneau) +
                             counter,
                             max_row=CoordinatesList[4],
                             min_col=22,
                             max_col=22)
        if result != None:
        ##print((CoordinatesList[0]-sizeofcreneau)+counter)
            while result != None and counter < sizeofcreneau:
                for Group, TypeSceance, Sceance, ProfName, Salle in Creneau1:
                    counter += 1
                    seances25.append({
                    'Group': Group.value,
                    'TypeSceance': TypeSceance.value,
                    'sceance': Sceance.value,
                    'ProfName': ProfName.value,
                    'Salle': Salle.value,
                })

        else:
            for temp in CreneauY1:
                print(counter)
                if countercours == 1:
                    seances25.append({'CasVide': temp[0].value})
                    counter += 1
                    countercours += 1
                elif countercours == 2:
                    seances25.append({
                    'NomMod': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 3:
                    seances25.append({
                    'Prof': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 4:
                    seances25.append({
                    'Salle_Cour': temp[0].value,
                })
                    counter += 1
                    countercours += 1

                elif countercours == 5:
                    seances25.append({'TypeSC': temp[0].value})
                    counter += 1
                    countercours += 1

                elif countercours == 6:
                    seances25.append({'CasVide2': temp[0].value})
                    counter += 1
                    countercours += 1
                elif counter > 6 and temp[0].value == None:
                    counter += 1
                    countercours += 1
                    if WS.cell(row=(CoordinatesList[4] - sizeofcreneau) + counter,
                           column=22).value is not None:
                        result = re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) +
                                    counter,
                                    column=22).value))
                    elif re.search(
                        "^G[1-9]$",
                        str(
                            WS.cell(row=(CoordinatesList[4] - sizeofcreneau) +
                                    counter,
                                    column=22).value)) is None:
                        countercours = 1
    return CoordinatesList,seances1,seances2,seances3,seances4,seances5,seances6,seances7,seances8,seances9,seances10,seances11,seances12,seances13,seances14,seances15,seances16,seances17,seances18,seances19,seances20,seances21,seances22,seances23,seances24,seances25



##inserts in the data base

##formation= input


def FillData(db,WS,FormationInput,NiveauInput,SemestreInput,AnnéeInput,CoordinatesList, seances1, seances2, seances3, seances4, seances5, seances6, seances7, seances8, seances9, seances10, seances11, seances12, seances13, seances14, seances15, seances16, seances17, seances18, seances19, seances20, seances21, seances22, seances23, seances24, seances25):
 db.insert({
        'formation': [{FormationInput,}],
        'Niveau': [{NiveauInput}],
        'Semestre':[{ SemestreInput}],
        'Année':[{AnnéeInput}],
        'Week': [{
            "dimanche" : [{
                'start': WS['A4'].value,
                'end': WS.cell(row=CoordinatesList[0], column=1).value,
                'sceances': seances1,
            }, {
                'start':
                WS.cell(row=CoordinatesList[0] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[1], column=1).value,
                'sceances':
                seances2,
            }, {
                'start':
                WS.cell(row=CoordinatesList[1] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[2], column=1).value,
                'sceances':
                seances3,
            }, {
                'start':
                WS.cell(row=CoordinatesList[2] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[3], column=1).value,
                'sceances':
                seances4,
            }, {
                'start':
                WS.cell(row=CoordinatesList[3] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[4], column=1).value,
                'sceances':
                seances5,
            }],
            'Lundi': [{
                'start': WS['A4'].value,
                'end': WS.cell(row=CoordinatesList[0], column=1).value,
                'sceances': seances6,
            }, {
                'start':
                WS.cell(row=CoordinatesList[0] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[1], column=1).value,
                'sceances':
                seances7,
            }, {
                'start':
                WS.cell(row=CoordinatesList[1] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[2], column=1).value,
                'sceances':
                seances8,
            }, {
                'start':
                WS.cell(row=CoordinatesList[2] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[3], column=1).value,
                'sceances':
                seances9,
            }, {
                'start':
                WS.cell(row=CoordinatesList[3] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[4], column=1).value,
                'sceances':
                seances10,
            }],
            'Mardi': [{
                'start': WS['A4'].value,
                'end': WS.cell(row=CoordinatesList[0], column=1).value,
                'sceances': seances11,
            }, {
                'start':
                WS.cell(row=CoordinatesList[0] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[1], column=1).value,
                'sceances':
                seances12,
            }, {
                'start':
                WS.cell(row=CoordinatesList[1] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[2], column=1).value,
                'sceances':
                seances13,
            }, {
                'start':
                WS.cell(row=CoordinatesList[2] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[3], column=1).value,
                'sceances':
                seances14,
            }, {
                'start':
                WS.cell(row=CoordinatesList[3] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[4], column=1).value,
                'sceances':
                seances15,
            }],
            'Mercredi': [{
                'start': WS['A4'].value,
                'end': WS.cell(row=CoordinatesList[0], column=1).value,
                'sceances': seances16,
            }, {
                'start':
                WS.cell(row=CoordinatesList[0] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[1], column=1).value,
                'sceances':
                seances17,
            }, {
                'start':
                WS.cell(row=CoordinatesList[1] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[2], column=1).value,
                'sceances':
                seances18,
            }, {
                'start':
                WS.cell(row=CoordinatesList[2] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[3], column=1).value,
                'sceances':
                seances19,
            }, {
                'start':
                WS.cell(row=CoordinatesList[3] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[4], column=1).value,
                'sceances':
                seances20,
            }],
            'Jeudi': [{
                'start': WS['A4'].value,
                'end': WS.cell(row=CoordinatesList[0], column=1).value,
                'sceances': seances21,
            }, {
                'start':
                WS.cell(row=CoordinatesList[0] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[1], column=1).value,
                'sceances':
                seances22,
            }, {
                'start':
                WS.cell(row=CoordinatesList[1] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[2], column=1).value,
                'sceances':
                seances23,
            }, {
                'start':
                WS.cell(row=CoordinatesList[2] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[3], column=1).value,
                'sceances':
                seances24,
            }, {
                'start':
                WS.cell(row=CoordinatesList[3] + 1, column=1).value,
                'end':
                WS.cell(row=CoordinatesList[4], column=1).value,
                'sceances':
                seances25,
            }],
        }]
    })




User = Query()



def printData(db):
 for item in db:
    print(item)

 db.all()





##Try fixing req3 4 6
##db.search(Req1.Week.any(Req2.dimanche.any(Query().start == "08H00")))
##db.search(Req1.Week.any(Req2.dimanche.any(Req3.Lundi.any(Req4.Mardi.any(Req5.Mercredi.any(Req6.Jeudi.any(Query().start == "08H00")))))))
def ResultFinder(db):
 Req1 = Query()
 Req2 = Query()
 Req3 = Query()
 Req4 = Query()
 Req5 = Query()
 Req6 = Query() 
 Result = db.search(Req1.Week.any(Req2.dimanche.any(Query().start == "08H00")))
 return Result
#(User.Week.Creneaux.start ==  '08H00')


def testField(value, test):
    return True if test == "" else value == test


def Search(Result,niveau, formation, semestre, Année, group, prof, salle):
    res1 = []
    res2 = []
    res3 = []
    res4 = []
    res5 = []
    ListFormation = []
    ListNiveau = []
    ListSemestre = []
    ListYear = []
    

    
    counter=['dimanche','Lundi','Mardi','Mercredi','Jeudi']
    for obj in Result:
     
      

     
     if testField(obj["formation"], formation) and testField(obj["Niveau"], niveau) and testField(obj["Semestre"], semestre): ##and testField(obj["Année"], Année):  #rajouter les tests sur l'année et le semestre
            for j in obj["Week"]:
                for cpt in range(0, 5):
                    for cr in j[counter[cpt]]:
                     for s in cr["sceances"]:
                      if "Group" in s:
                        if testField(s["Group"], group) and testField(s["ProfName"], prof) and testField(s["Salle"], salle):
                            if cpt==0:
                             res1.append(s)
                             print(s)
                            elif cpt==1:
                             res2.append(s)
                             print(s)
                            elif cpt==2:
                             res3.append(s)
                             print(s)
                            elif cpt==3:
                             res4.append(s)
                             print(s)
                            else:
                             res5.append(s)
                             print(s)
    
    for cpt in range(0, 4):
        for item in obj:
            if cpt == 0:
                ListFormation.append(item)
            elif cpt == 1:
                ListNiveau.append(item)
            elif cpt == 2:
                ListSemestre.append(item)
            elif cpt==3:
                ListYear.append(item)

    print(ListFormation, ListNiveau ,ListSemestre , ListYear)



    return res1,res2,res3,res4,res5
def EmptyDB(db):
 db.truncate()
