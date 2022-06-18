from asyncio import subprocess
from fileinput import filename
from importlib.resources import path
from msilib.schema import Directory
import os
from sqlite3 import dbapi2
import sys

from tkinter import dialog
from PyQt5 import QtCore, QtGui, QtWidgets, QtQuickWidgets
from PyQt5.QtWidgets import QApplication, QComboBox, QPushButton, QFileDialog, QVBoxLayout, QLineEdit
from pyparsing import empty
from EmploiApp import Ui_MainWindow
import os.path
from tinydb import TinyDB, Query
from jinja2 import Environment, FileSystemLoader, select_autoescape
from openpyxl import Workbook, load_workbook
import webbrowser
from DBtest import DBcreate, EmptyDB, FillData, LoadXLFile, ResultFinder, Search, parcourirCreneau, printData
import subprocess

##cmd='python test_jinja.py'


class MyMainWindow (Ui_MainWindow):
    def __init__(self, MW) -> None:
        super().__init__()
        self.setupUi(MW)
        self.MW = MW
        # Put your events here
        # Buttons
        self.ButtonImport.clicked.connect(self.HideFrame)
        self.ButtonAdd.clicked.connect(self.select)
        self.RemplirDB.clicked.connect(self.ShowFormat)
        self.ButtonBack.clicked.connect(self.ShowFrame)
        self.ButtonCreate.clicked.connect(self.SearchAndHTML)
        self.ButtonEmpty.clicked.connect(self.Empty)

  # Define the events here
    def ShowFormat(self):
        os.startfile("TestPic.jpg")

    def Empty(self):
        db = DBcreate()
        EmptyDB(db)


    def HideFrame(self):
        self.frame.hide()

    def ShowFrame(self):
        self.frame.show()

    def AddInfo(self):
        self.niveau = self.Niveau.text()
        self.formation = self.Formation.text()
        self.semestre = self.Semestre.text()
        self.Année = self.Year.text()
        FillData(self.formation, self.niveau, self.semestre, self.Année)

        self.frame.show()
        # Search(self.niveau,self.formation,self.semestre,self.Année,self.Group,self.Prof,self.Salle)

    def SearchAndHTML(self):

        self.TeacherSearch = self.Prof.text()
        self.GroupSearch = self.Group.text()
        self.SalleSearch = self.Salle.text()
        self.niveau = self.Niveau.text()
        self.formation = self.Formation.text()
        self.semestre = self.Semestre.text()
        self.Année = self.Year.text()
        db = DBcreate()
        Result = ResultFinder(db)
        [res1, res2, res3, res4, res5,StartsAndEnds] = Search(
            Result, self.niveau, self.formation, self.semestre, self.Année, self.GroupSearch, self.TeacherSearch, self.SalleSearch)
        
        print(len(res1))
        
        firstD="N/A"
        firstL="N/A"
        firstM="N/A"
        firstMR="N/A"
        firstJ="N/A"
        SecondD="N/A"
        SecondL="N/A"
        SecondM="N/A"
        SecondMR="N/A"
        SecondJ="N/A"
        ThirdD="N/A"
        ThirdL="N/A"
        ThirdM="N/A"
        ThirdMR="N/A"
        ThirdJ="N/A"
        FourthD="N/A"
        FourthL="N/A"
        FourthM="N/A"
        FourthMR="N/A"
        FourthJ="N/A"
        FifthD="N/A"
        FifthL="N/A"
        FifthM="N/A"
        FifthMR="N/A"
        FifthJ="N/A"
        ##Dimanche
        for cpt in range(0,5):
            for counter in range(0,len(res1)):
               if bool(res1[counter])==True:
                   for s in res1[counter]:
                       if "Group" in s and res1[counter]["TypeSceance"] is not None:
                           if res1[counter]["Begins"] == StartsAndEnds[cpt*2]:
                               if cpt==0:
                                   firstD=res1[counter]["Group"]+" Seance "+res1[counter]["TypeSceance"]+" "+res1[counter]["sceance"]+" Avec Prof "+res1[counter]["ProfName"]+" à salle "+res1[counter]["Salle"]
                               elif cpt==1:
                                   SecondD=res1[counter]["Group"]+" Seance "+res1[counter]["TypeSceance"]+" "+res1[counter]["sceance"]+" Avec Prof "+res1[counter]["ProfName"]+" à salle "+res1[counter]["Salle"]
                               elif cpt==2:
                                   ThirdD=res1[counter]["Group"]+" Seance "+res1[counter]["TypeSceance"]+" "+res1[counter]["sceance"]+" Avec Prof "+res1[counter]["ProfName"]+" à salle "+res1[counter]["Salle"]
                               elif cpt==3:
                                   FourthD=res1[counter]["Group"]+" Seance "+res1[counter]["TypeSceance"]+" "+res1[counter]["sceance"]+" Avec Prof "+res1[counter]["ProfName"]+" à salle "+res1[counter]["Salle"]
                               elif cpt==4:
                                   FifthD=res1[counter]["Group"]+" Seance "+res1[counter]["TypeSceance"]+" "+res1[counter]["sceance"]+" Avec Prof "+res1[counter]["ProfName"]+" à salle "+res1[counter]["Salle"]
                           
                       elif "NomMod" in s:
                           if res1[counter]["Begins"] == StartsAndEnds[cpt*2]:
                               if cpt==0:
                                   firstD=res1[counter]["TypeSC"]+" "+res1[counter]["NomMod"]+" Avec "+res1[counter]["Prof"]+" Dans Salle "+res1[counter]["Salle_Cour"]
                               elif cpt==1:
                                   SecondD=res1[counter]["TypeSC"]+" "+res1[counter]["NomMod"]+" Avec "+res1[counter]["Prof"]+" Dans Salle "+res1[counter]["Salle_Cour"]
                               elif cpt==2:
                                   ThirdD=res1[counter]["TypeSC"]+" "+res1[counter]["NomMod"]+" Avec "+res1[counter]["Prof"]+" Dans Salle "+res1[counter]["Salle_Cour"]
                               elif cpt==3:
                                   FourthD=res1[counter]["TypeSC"]+" "+res1[counter]["NomMod"]+" Avec "+res1[counter]["Prof"]+" Dans Salle "+res1[counter]["Salle_Cour"]
                               elif cpt==4:
                                   FifthD=res1[counter]["TypeSC"]+" "+res1[counter]["NomMod"]+" Avec "+res1[counter]["Prof"]+" Dans Salle "+res1[counter]["Salle_Cour"]
                                  

         
        
        ##Lundi 
        
        for cpt in range(0,5):
            for counter in range(0,len(res2)):
               if bool(res2[counter])==True:
                   for s in res2[counter]:
                       if "Group" in s and res2[counter]["TypeSceance"] is not None:
                           if res2[counter]["Begins"] == StartsAndEnds[cpt*2]:
                               if cpt==0:
                                   firstL=res2[counter]["Group"]+" Seance "+res2[counter]["TypeSceance"]+" "+res2[counter]["sceance"]+" Avec Prof "+res2[counter]["ProfName"]+" à salle "+res2[counter]["Salle"]
                               elif cpt==1:
                                   SecondL=res2[counter]["Group"]+" Seance "+res2[counter]["TypeSceance"]+" "+res2[counter]["sceance"]+" Avec Prof "+res2[counter]["ProfName"]+" à salle "+res2[counter]["Salle"]
                               elif cpt==2:
                                   ThirdL=res2[counter]["Group"]+" Seance "+res2[counter]["TypeSceance"]+" "+res2[counter]["sceance"]+" Avec Prof "+res2[counter]["ProfName"]+" à salle "+res2[counter]["Salle"]
                               elif cpt==3:
                                   FourthL=res2[counter]["Group"]+" Seance "+res2[counter]["TypeSceance"]+" "+res2[counter]["sceance"]+" Avec Prof "+res2[counter]["ProfName"]+" à salle "+res2[counter]["Salle"]
                               elif cpt==4:
                                   FifthL=res2[counter]["Group"]+" Seance "+res2[counter]["TypeSceance"]+" "+res2[counter]["sceance"]+" Avec Prof "+res2[counter]["ProfName"]+" à salle "+res2[counter]["Salle"]
                           
                       elif "NomMod" in s:
                           if res2[counter]["Begins"] == StartsAndEnds[cpt*2]:
                               if cpt==0:
                                   firstL=res2[counter]["TypeSC"]+" "+res2[counter]["NomMod"]+" Avec "+res2[counter]["Prof"]+" Dans Salle "+res2[counter]["Salle_Cour"]
                               elif cpt==1:
                                   SecondL=res2[counter]["TypeSC"]+" "+res2[counter]["NomMod"]+" Avec "+res2[counter]["Prof"]+" Dans Salle "+res2[counter]["Salle_Cour"]
                               elif cpt==2:
                                   ThirdL=res2[counter]["TypeSC"]+" "+res2[counter]["NomMod"]+" Avec "+res2[counter]["Prof"]+" Dans Salle "+res2[counter]["Salle_Cour"]
                               elif cpt==3:
                                   FourthL=res2[counter]["TypeSC"]+" "+res2[counter]["NomMod"]+" Avec "+res2[counter]["Prof"]+" Dans Salle "+res2[counter]["Salle_Cour"]
                               elif cpt==4:
                                   FifthL=res2[counter]["TypeSC"]+" "+res2[counter]["NomMod"]+" Avec "+res2[counter]["Prof"]+" Dans Salle "+res2[counter]["Salle_Cour"]
                           
                                    
                                 

        
        ##Mardi
        for cpt in range(0,5):
            for counter in range(0,len(res3)):
               if bool(res3[counter])==True:
                   for s in res3[counter]:
                       if "Group" in s and res3[counter]["TypeSceance"] is not None:
                           if res3[counter]["Begins"] == StartsAndEnds[cpt*2]:
                               if cpt==0:
                                   firstM=res3[counter]["Group"]+" Seance "+res3[counter]["TypeSceance"]+" "+res3[counter]["sceance"]+" Avec Prof "+res3[counter]["ProfName"]+" à salle "+res3[counter]["Salle"]
                               elif cpt==1:
                                   SecondM=res3[counter]["Group"]+" Seance "+res3[counter]["TypeSceance"]+" "+res3[counter]["sceance"]+" Avec Prof "+res3[counter]["ProfName"]+" à salle "+res3[counter]["Salle"]
                               elif cpt==2:
                                   ThirdM=res3[counter]["Group"]+" Seance "+res3[counter]["TypeSceance"]+" "+res3[counter]["sceance"]+" Avec Prof "+res3[counter]["ProfName"]+" à salle "+res3[counter]["Salle"]
                               elif cpt==3:
                                   FourthM=res3[counter]["Group"]+" Seance "+res3[counter]["TypeSceance"]+" "+res3[counter]["sceance"]+" Avec Prof "+res3[counter]["ProfName"]+" à salle "+res3[counter]["Salle"]
                               elif cpt==4:
                                   FifthM=res3[counter]["Group"]+" Seance "+res3[counter]["TypeSceance"]+" "+res3[counter]["sceance"]+" Avec Prof "+res3[counter]["ProfName"]+" à salle "+res3[counter]["Salle"]
                           
                       elif "NomMod" in s:
                           if res3[counter]["Begins"] == StartsAndEnds[cpt*2]:
                               if cpt==0:
                                   firstM=res3[counter]["TypeSC"]+" "+res3[counter]["NomMod"]+" Avec "+res3[counter]["Prof"]+" Dans Salle "+res3[counter]["Salle_Cour"]
                               elif cpt==1:
                                   SecondM=res3[counter]["TypeSC"]+" "+res3[counter]["NomMod"]+" Avec "+res3[counter]["Prof"]+" Dans Salle "+res3[counter]["Salle_Cour"]
                               elif cpt==2:
                                   ThirdM=res3[counter]["TypeSC"]+" "+res3[counter]["NomMod"]+" Avec "+res3[counter]["Prof"]+" Dans Salle "+res3[counter]["Salle_Cour"]
                               elif cpt==3:
                                   FourthM=res3[counter]["TypeSC"]+" "+res3[counter]["NomMod"]+" Avec "+res3[counter]["Prof"]+" Dans Salle "+res3[counter]["Salle_Cour"]
                               elif cpt==4:
                                   FifthM=res3[counter]["TypeSC"]+" "+res3[counter]["NomMod"]+" Avec "+res3[counter]["Prof"]+" Dans Salle "+res3[counter]["Salle_Cour"]
                            
                        
                  
        ##Mercredi

        for cpt in range(0,5):
            for counter in range(0,len(res4)):
               if bool(res4[counter])==True:
                   for s in res4[counter]:
                       if "Group" in s and res4[counter]["TypeSceance"] is not None:
                           if res4[counter]["Begins"] == StartsAndEnds[cpt*2]:
                               if cpt==0:
                                   firstMR=res4[counter]["Group"]+" Seance "+res4[counter]["TypeSceance"]+" "+res4[counter]["sceance"]+" Avec Prof "+res4[counter]["ProfName"]+" à salle "+res4[counter]["Salle"]
                               elif cpt==1:
                                   SecondMR=res4[counter]["Group"]+" Seance "+res4[counter]["TypeSceance"]+" "+res4[counter]["sceance"]+" Avec Prof "+res4[counter]["ProfName"]+" à salle "+res4[counter]["Salle"]
                               elif cpt==2:
                                   ThirdMR=res4[counter]["Group"]+" Seance "+res4[counter]["TypeSceance"]+" "+res4[counter]["sceance"]+" Avec Prof "+res4[counter]["ProfName"]+" à salle "+res4[counter]["Salle"]
                               elif cpt==3:
                                   FourthMR=res4[counter]["Group"]+" Seance "+res4[counter]["TypeSceance"]+" "+res4[counter]["sceance"]+" Avec Prof "+res4[counter]["ProfName"]+" à salle "+res4[counter]["Salle"]
                               elif cpt==4:
                                   FifthMR=res4[counter]["Group"]+" Seance "+res4[counter]["TypeSceance"]+" "+res4[counter]["sceance"]+" Avec Prof "+res4[counter]["ProfName"]+" à salle "+res4[counter]["Salle"]
                           else:
                               if cpt==0:
                                   firstMR="N/A"
                               elif cpt==1:
                                   SecondMR="N/A"
                               elif cpt==2:
                                   ThirdMR="N/A"
                               elif cpt==3:
                                   FourthMR="N/A"
                               elif cpt==4:
                                   FifthMR="N/A"
                       elif "NomMod" in s:
                           if res4[counter]["Begins"] == StartsAndEnds[cpt*2]:
                               if cpt==0:
                                   firstMR=res4[counter]["TypeSC"]+" "+res4[counter]["NomMod"]+" Avec "+res4[counter]["Prof"]+" Dans Salle "+res4[counter]["Salle_Cour"]
                               elif cpt==1:
                                   SecondMR=res4[counter]["TypeSC"]+" "+res4[counter]["NomMod"]+" Avec "+res4[counter]["Prof"]+" Dans Salle "+res4[counter]["Salle_Cour"]
                               elif cpt==2:
                                   ThirdMR=res4[counter]["TypeSC"]+" "+res4[counter]["NomMod"]+" Avec "+res4[counter]["Prof"]+" Dans Salle "+res4[counter]["Salle_Cour"]
                               elif cpt==3:
                                   FourthMR=res4[counter]["TypeSC"]+" "+res4[counter]["NomMod"]+" Avec "+res4[counter]["Prof"]+" Dans Salle "+res4[counter]["Salle_Cour"]
                               elif cpt==4:
                                   FifthMR=res4[counter]["TypeSC"]+" "+res4[counter]["NomMod"]+" Avec "+res4[counter]["Prof"]+" Dans Salle "+res4[counter]["Salle_Cour"]
                           
                       
               
        
        ## Jeudi
        for cpt in range(0,5):
            for counter in range(0,len(res5)):
               if bool(res5[counter])==True:
                   for s in res5[counter]:
                       if "Group" in s and res5[counter]["TypeSceance"] is not None:
                           if res5[counter]["Begins"] == StartsAndEnds[cpt*2]:
                               if cpt==0:
                                   firstJ=res5[counter]["Group"]+" Seance "+res5[counter]["TypeSceance"]+" "+res5[counter]["sceance"]+" Avec Prof "+res5[counter]["ProfName"]+" à salle "+res5[counter]["Salle"]
                               elif cpt==1:
                                   SecondJ=res5[counter]["Group"]+" Seance "+res5[counter]["TypeSceance"]+" "+res5[counter]["sceance"]+" Avec Prof "+res5[counter]["ProfName"]+" à salle "+res5[counter]["Salle"]
                               elif cpt==2:
                                   ThirdJ=res5[counter]["Group"]+" Seance "+res5[counter]["TypeSceance"]+" "+res5[counter]["sceance"]+" Avec Prof "+res5[counter]["ProfName"]+" à salle "+res5[counter]["Salle"]
                               elif cpt==3:
                                   FourthJ=res5[counter]["Group"]+" Seance "+res5[counter]["TypeSceance"]+" "+res5[counter]["sceance"]+" Avec Prof "+res5[counter]["ProfName"]+" à salle "+res5[counter]["Salle"]
                               elif cpt==4:
                                   FifthJ=res5[counter]["Group"]+" Seance "+res5[counter]["TypeSceance"]+" "+res5[counter]["sceance"]+" Avec Prof "+res5[counter]["ProfName"]+" à salle "+res5[counter]["Salle"]
                           
                       elif "NomMod" in s:
                           if res5[cpt]["Begins"] == StartsAndEnds[cpt*2]:
                               if cpt==0:
                                   firstJ=res5[cpt]["TypeSC"]+" "+res5[cpt]["NomMod"]+" Avec "+res5[cpt]["Prof"]+" Dans Salle "+res5[cpt]["Salle_Cour"]
                               elif cpt==1:
                                   SecondJ=res5[cpt]["TypeSC"]+" "+res5[cpt]["NomMod"]+" Avec "+res5[cpt]["Prof"]+" Dans Salle "+res5[cpt]["Salle_Cour"]
                               elif cpt==2:
                                   ThirdJ=res5[cpt]["TypeSC"]+" "+res5[cpt]["NomMod"]+" Avec "+res5[cpt]["Prof"]+" Dans Salle "+res5[cpt]["Salle_Cour"]
                               elif cpt==3:
                                   FourthJ=res5[cpt]["TypeSC"]+" "+res5[cpt]["NomMod"]+" Avec "+res5[cpt]["Prof"]+" Dans Salle "+res5[cpt]["Salle_Cour"]
                               elif cpt==4:
                                   FifthJ=res5[cpt]["TypeSC"]+" "+res5[cpt]["NomMod"]+" Avec "+res5[cpt]["Prof"]+" Dans Salle "+res5[cpt]["Salle_Cour"]
                           

                  

               
            
       
        

        
        """
        if res1[1]["Begins"] == StartsAndEnds[2]:
         SecondD= res1[1]["Group"]+" Seance "+res1[1]["TypeSceance"]+" "+res1[1]["sceance"]+" Avec Prof "+res1[1]["ProfName"]+" à salle "+res1[1]["Salle"]
        else:
            SecondD="N/A"
        """
        # Jinja2 part
        env = Environment(
            loader=FileSystemLoader("./"),
            autoescape=select_autoescape()
        )
        
        template = env.get_template("templ.html")
        output = open("emploi.html", "w")

        output.write(
            template.render(
                formation=self.formation,
                niveau=self.niveau,
                semestre=self.semestre,
                emploi=[
                    {"creneau":StartsAndEnds[0]+"-"+StartsAndEnds[1],
                    "dimanche":firstD,
                    "lundi":firstL,
                    "mardi":firstM,
                    "mercredi":firstMR, 
                    "jeudi":firstJ},
                    {"creneau": StartsAndEnds[2]+"-"+StartsAndEnds[3], 
                    "dimanche": SecondD,
                    "lundi":SecondL,
                    "mardi":SecondM, 
                    "mercredi":SecondMR, 
                    "jeudi":SecondJ},
                    {"creneau":StartsAndEnds[4]+"-"+StartsAndEnds[5],
                    "dimanche":ThirdD,
                    "lundi":ThirdL,
                    "mardi":ThirdM,
                    "mercredi":ThirdMR,
                    "jeudi":ThirdJ},
                    {"creneau":StartsAndEnds[6]+"-"+StartsAndEnds[7],
                    "dimanche":FourthD,
                    "lundi":FourthL,
                    "mardi":FourthM,
                    "mercredi":FourthMR,
                    "jeudi":FourthJ},
                    {"creneau":StartsAndEnds[8]+"-"+StartsAndEnds[9],
                    "dimanche":FifthD,
                    "lundi":FifthL,
                    "mardi":FifthM,
                    "mercredi":FifthMR,
                    "jeudi":FifthJ},

                ]))
        output.close()

        webbrowser.open("file://" + os.path.join(os.getcwd(), "emploi.html"))
        EmptyDB(db)

    def select(self):

        #response =  QFileDialog.getOpenFileName(   parent=self,caption='Select file',Directory=os.getcwd(), filter=file_filter,initialFilter='Excel File (*.xlsx *.xls)')
        # return response [0]

     ##path ,ext = QtWidgets.QFileDialog.getOpenFileName(self.MW, 'Select a file',filter='Data file (*.xlsx *.csv .dat);; Excel File (*.xlsx *.xls)')
     # os.path.exists(ext)
     # print(str(ext))

        FileName, _ = QFileDialog.getOpenFileName(
            self.MW, 'Ouvrir Fichier', 'C:\ ', "xlsx files (*.xlsx)")
        if not FileName:
            return
        print(FileName)
        path = FileName
        WS = LoadXLFile(path)
        parcourirCreneau(WS)
        CoordinatesList, seances1, seances2, seances3, seances4, seances5, seances6, seances7, seances8, seances9, seances10, seances11, seances12, seances13, seances14, seances15, seances16, seances17, seances18, seances19, seances20, seances21, seances22, seances23, seances24, seances25 = parcourirCreneau(
            WS)
        self.niveau = self.Niveau.text()
        self.formation = self.Formation.text()
        self.semestre = self.Semestre.text()
        self.Année = self.Year.text()

        db = DBcreate()
        FillData(db, WS, self.formation, self.niveau, self.semestre, self.semestre, CoordinatesList, seances1, seances2, seances3, seances4, seances5, seances6, seances7, seances8, seances9,
                 seances10, seances11, seances12, seances13, seances14, seances15, seances16, seances17, seances18, seances19, seances20, seances21, seances22, seances23, seances24, seances25)

        printData(db)


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = MyMainWindow(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
