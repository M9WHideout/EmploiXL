from asyncio import subprocess
from fileinput import filename
from importlib.resources import path
from msilib.schema import Directory
import os
from sqlite3 import dbapi2
import sys
from tkinter import dialog
from PyQt5 import QtCore, QtGui, QtWidgets, QtQuickWidgets 
from PyQt5.QtWidgets import QApplication, QComboBox,QPushButton, QFileDialog, QVBoxLayout,QLineEdit
from pyparsing import empty
from EmploiApp import Ui_MainWindow
import  os.path
from tinydb import TinyDB,Query
from jinja2 import Environment,FileSystemLoader, select_autoescape
from openpyxl import Workbook, load_workbook 
import webbrowser
from DBtest import DBcreate, EmptyDB, FillData, LoadXLFile, ResultFinder, Search, parcourirCreneau, printData
import subprocess
##cmd='python test_jinja.py'
class MyMainWindow (Ui_MainWindow):
    def __init__(self,MW) -> None:
        super().__init__()
        self.setupUi(MW)
        self.MW=MW
        # Put your events here
        self.ButtonImport.clicked.connect(self.HideFrame)
        self.ButtonAdd.clicked.connect(self.select)
        self.RemplirDB.clicked.connect(self.HideFrame)
        self.ButtonBack.clicked.connect(self.ShowFrame)
        self.ButtonCreate.clicked.connect(self.SearchAndHTML)
        
    
  # Define the events here

    def HideFrame(self):
        self.frame.hide()

    def ShowFrame(self):
        self.frame.show() 


    def AddInfo(self):
        self.niveau = self.Niveau.text()
        self.formation = self.Formation.text()
        self.semestre=self.Semestre.text()
        self.Année=self.Year.text()
        FillData(self.formation,self.niveau,self.semestre,self.Année)

        self.frame.show()
        ##Search(self.niveau,self.formation,self.semestre,self.Année,self.Group,self.Prof,self.Salle)
        
 
    def SearchAndHTML(self):

        self.TeacherSearch = self.Prof.text()
        self.GroupSearch = self.Group.text()
        self.SalleSearch = self.Salle.text()
        self.niveau = self.Niveau.text()
        self.formation = self.Formation.text()
        self.semestre=self.Semestre.text()
        self.Année=self.Year.text()
        db=DBcreate()
        Result=ResultFinder(db)
        res=Search(Result,self.niveau, self.formation, self.semestre, self.Année, self.GroupSearch, self.TeacherSearch, self.SalleSearch)
        print(res)


        ##Jinja2 part
        env = Environment(
        loader=FileSystemLoader("./"),
        autoescape=select_autoescape()
        ) 
        template = env.get_template("templ.html")
        output=open("emploi.html","w")
          
        output.write(
        template.render(
        a_variable="", 
        niveau="Niveau L2",
        emploi=[
            {"creneau":"0sss8H00-09H00","dimanche":"G1', 'TypeSceance': 'TD', 'sceance': 'IHM', 'ProfName': 'Kahya', 'Salle': 'AG51","lundi":"G1', 'TypeSceance': 'TD', 'sceance': 'Comp', 'ProfName': 'Amrane', 'Salle': 'INF2","mardi":"G1', 'TypeSceance': 'Cours', 'sceance': 'PS', 'ProfName': 'Redjil', 'Salle': 'A11","mercredi":"G1', 'TypeSceance': 'TD', 'sceance': 'IHM', 'ProfName': 'Kahya', 'Salle': 'AG51","jeudi":"G1', 'TypeSceance': 'TP', 'sceance': 'SE', 'ProfName': 'Hariati', 'Salle': 'S1"},
            {"creneau":"09H15-10H15","dimanche":"G1', 'TypeSceance': 'TP', 'sceance': 'IHM', 'ProfName': 'Kahya', 'Salle': 'S1","lundi":"G1', 'TypeSceance': 'TD', 'sceance': 'PS', 'ProfName': 'Redjili', 'Salle': 'AG51","mardi":"G1', 'TypeSceance': 'TD', 'sceance': 'IHM', 'ProfName': 'Kahya', 'Salle': 'AG51","mercredi":"G1', 'TypeSceance': 'TD', 'sceance': 'Comp', 'ProfName': 'Sari', 'Salle': 'AG51","jeudi":"G1', 'TypeSceance': 'TP', 'sceance': 'GL', 'ProfName': 'Atil', 'Salle': 'S1"},
            ##{"creneau":"","dimanche":"Seance 2 du jour 1","lundi":"Seance 2 du jour 2","mardi":"Seance 2 du jour 3","mercredi":"Seance 2 du jour 4","jeudi":"Seance 2 du jour 5"},
            
        ]))
        output.close()
       
       








        
        webbrowser.open("file://" + os.path.join(os.getcwd(), "emploi.html"))
        EmptyDB(db)
        return res
      
   
    def select(self):
        
        #response =  QFileDialog.getOpenFileName(   parent=self,caption='Select file',Directory=os.getcwd(), filter=file_filter,initialFilter='Excel File (*.xlsx *.xls)')
        #return response [0]
    
     ##path ,ext = QtWidgets.QFileDialog.getOpenFileName(self.MW, 'Select a file',filter='Data file (*.xlsx *.csv .dat);; Excel File (*.xlsx *.xls)')
     ##os.path.exists(ext)
     ##print(str(ext))
      
      FileName,_=QFileDialog.getOpenFileName(self.MW,'Ouvrir Fichier', 'C:\ ' ,"xlsx files (*.xlsx)")
      if not FileName:return
      print(FileName)
      path=FileName
      WS=LoadXLFile(path)
      parcourirCreneau(WS)
      CoordinatesList, seances1, seances2, seances3, seances4, seances5, seances6, seances7, seances8, seances9, seances10, seances11, seances12, seances13, seances14, seances15, seances16, seances17, seances18, seances19, seances20, seances21, seances22, seances23, seances24, seances25 = parcourirCreneau(WS)
      self.niveau = self.Niveau.text()
      self.formation = self.Formation.text()
      self.semestre=self.Semestre.text()
      self.Année=self.Year.text()
      

      db=DBcreate()
      FillData(db,WS,self.formation,self.niveau,self.semestre,self.semestre,CoordinatesList, seances1, seances2, seances3, seances4, seances5, seances6, seances7, seances8, seances9, seances10, seances11, seances12, seances13, seances14, seances15, seances16, seances17, seances18, seances19, seances20, seances21, seances22, seances23, seances24, seances25)
      
      printData(db)




      
      

    





      
        
if __name__ == "__main__":
    import sys    
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = MyMainWindow(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())