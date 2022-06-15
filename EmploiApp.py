# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'EmploiApp.ui'
#
# Created by: PyQt5 UI code generator 5.15.6
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(538, 463)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(10, 10, 551, 451))
        self.frame_2.setStyleSheet("background-color: rgb(61, 80, 109);")
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.label_5 = QtWidgets.QLabel(self.frame_2)
        self.label_5.setGeometry(QtCore.QRect(80, 240, 111, 52))
        font = QtGui.QFont()
        font.setFamily("Segoe UI Black")
        font.setPointSize(12)
        font.setBold(True)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.Formation = QtWidgets.QLineEdit(self.frame_2)
        self.Formation.setGeometry(QtCore.QRect(200, 250, 132, 30))
        self.Formation.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"\n"
"border: none;\n"
"border-radius: 6px;")
        self.Formation.setObjectName("Formation")
        self.ButtonBack = QtWidgets.QPushButton(self.frame_2)
        self.ButtonBack.setGeometry(QtCore.QRect(214, 373, 91, 31))
        self.ButtonBack.setStyleSheet("QPushButton{\n"
"   background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #71b7e6, stop:1 #9b60b6);\n"
"    \n"
"    border: none;\n"
"    border-radius:  10px;\n"
"    color: #FFF\n"
"}\n"
"")
        self.ButtonBack.setObjectName("ButtonBack")
        self.ButtonAdd = QtWidgets.QPushButton(self.frame_2)
        self.ButtonAdd.setGeometry(QtCore.QRect(214, 333, 91, 31))
        self.ButtonAdd.setStyleSheet("QPushButton{\n"
"\n"
"       background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #71b7e6, stop:1 #9b60b6);\n"
"    border: none;\n"
"    border-radius:  10px;\n"
"    color: #FFF\n"
"}")
        self.ButtonAdd.setObjectName("ButtonAdd")
        self.label_6 = QtWidgets.QLabel(self.frame_2)
        self.label_6.setGeometry(QtCore.QRect(80, 290, 81, 21))
        font = QtGui.QFont()
        font.setFamily("Segoe UI Black")
        font.setPointSize(12)
        font.setBold(True)
        self.label_6.setFont(font)
        self.label_6.setStyleSheet("family-font: \'Poppins\', sans-serif;")
        self.label_6.setObjectName("label_6")
        self.label_7 = QtWidgets.QLabel(self.frame_2)
        self.label_7.setGeometry(QtCore.QRect(80, 200, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Segoe UI Black")
        font.setPointSize(12)
        font.setBold(True)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.Niveau = QtWidgets.QLineEdit(self.frame_2)
        self.Niveau.setGeometry(QtCore.QRect(200, 290, 132, 30))
        self.Niveau.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"\n"
"border: none;\n"
"border-radius: 6px;")
        self.Niveau.setObjectName("Niveau")
        self.Year = QtWidgets.QLineEdit(self.frame_2)
        self.Year.setGeometry(QtCore.QRect(200, 200, 132, 30))
        self.Year.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"border: none;\n"
"border-radius: 6px;")
        self.Year.setObjectName("Year")
        self.label_8 = QtWidgets.QLabel(self.frame_2)
        self.label_8.setGeometry(QtCore.QRect(80, 160, 131, 31))
        font = QtGui.QFont()
        font.setFamily("Segoe UI Black")
        font.setPointSize(12)
        font.setBold(True)
        self.label_8.setFont(font)
        self.label_8.setObjectName("label_8")
        self.Semestre = QtWidgets.QLineEdit(self.frame_2)
        self.Semestre.setGeometry(QtCore.QRect(200, 160, 132, 30))
        self.Semestre.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"border: none;\n"
"border-radius: 6px;")
        self.Semestre.setObjectName("Semestre")
        self.Warning2 = QtWidgets.QLabel(self.frame_2)
        self.Warning2.setGeometry(QtCore.QRect(90, 60, 491, 71))
        font = QtGui.QFont()
        font.setFamily("Segoe UI Black")
        font.setPointSize(12)
        font.setBold(True)
        self.Warning2.setFont(font)
        self.Warning2.setStyleSheet("color: rgb(255, 0, 0);")
        self.Warning2.setObjectName("Warning2")
        self.frame = QtWidgets.QFrame(self.frame_2)
        self.frame.setGeometry(QtCore.QRect(-30, -10, 801, 581))
        self.frame.setStyleSheet("background-color: rgb(61, 80, 109);")
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.Group = QtWidgets.QLineEdit(self.frame)
        self.Group.setGeometry(QtCore.QRect(280, 110, 113, 31))
        self.Group.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"border: none;\n"
"border-radius: 6px;")
        self.Group.setObjectName("Group")
        self.ButtonCreate = QtWidgets.QPushButton(self.frame)
        self.ButtonCreate.setGeometry(QtCore.QRect(210, 340, 101, 31))
        self.ButtonCreate.setStyleSheet("QPushButton{\n"
"   background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #71b7e6, stop:1 #9b60b6);\n"
"    \n"
"    border: none;\n"
"    border-radius:  10px;\n"
"    color: #FFF\n"
"}\n"
"")
        self.ButtonCreate.setObjectName("ButtonCreate")
        self.ButtonImport = QtWidgets.QPushButton(self.frame)
        self.ButtonImport.setGeometry(QtCore.QRect(270, 290, 121, 31))
        self.ButtonImport.setStyleSheet("QPushButton{\n"
"\n"
"       background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #71b7e6, stop:1 #9b60b6);\n"
"    border: none;\n"
"    border-radius:  10px;\n"
"    color: #FFF\n"
"}")
        self.ButtonImport.setObjectName("ButtonImport")
        self.Prof = QtWidgets.QLineEdit(self.frame)
        self.Prof.setGeometry(QtCore.QRect(150, 110, 111, 31))
        self.Prof.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"border: none;\n"
"border-radius: 6px;")
        self.Prof.setObjectName("Prof")
        self.Salle = QtWidgets.QLineEdit(self.frame)
        self.Salle.setGeometry(QtCore.QRect(210, 150, 113, 31))
        self.Salle.setStyleSheet("background-color: rgb(255, 255, 255);\n"
"border: none;\n"
"border-radius: 6px;")
        self.Salle.setObjectName("Salle")
        self.RemplirDB = QtWidgets.QPushButton(self.frame)
        self.RemplirDB.setGeometry(QtCore.QRect(140, 290, 101, 31))
        self.RemplirDB.setStyleSheet("QPushButton{\n"
"   background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #71b7e6, stop:1 #9b60b6);\n"
"    \n"
"    border: none;\n"
"    border-radius:  10px;\n"
"    color: #FFF\n"
"}\n"
"")
        self.RemplirDB.setObjectName("RemplirDB")
        self.Warning1 = QtWidgets.QLabel(self.frame)
        self.Warning1.setGeometry(QtCore.QRect(100, 10, 351, 91))
        font = QtGui.QFont()
        font.setFamily("Segoe UI Black")
        font.setPointSize(12)
        font.setBold(True)
        self.Warning1.setFont(font)
        self.Warning1.setStyleSheet("color: rgb(255, 0, 0);")
        self.Warning1.setObjectName("Warning1")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label_5.setText(_translate("MainWindow", "Formation:"))
        self.ButtonBack.setText(_translate("MainWindow", "Retour"))
        self.ButtonAdd.setText(_translate("MainWindow", "Import fichier"))
        self.label_6.setText(_translate("MainWindow", "Niveau:"))
        self.label_7.setText(_translate("MainWindow", "*Année:"))
        self.label_8.setText(_translate("MainWindow", "*Semestre:"))
        self.Warning2.setText(_translate("MainWindow", "Veuillez remplir toutes les informations \n"
"                         * nécessaires"))
        self.Group.setPlaceholderText(_translate("MainWindow", "Group ici"))
        self.ButtonCreate.setText(_translate("MainWindow", "Créer"))
        self.ButtonImport.setText(_translate("MainWindow", "Import Fichier"))
        self.Prof.setPlaceholderText(_translate("MainWindow", "Nom enseignant"))
        self.Salle.setPlaceholderText(_translate("MainWindow", "Salle ici"))
        self.RemplirDB.setText(_translate("MainWindow", "Remplir le \n"
"formulaire"))
        self.Warning1.setText(_translate("MainWindow", "ajoutez des mots clés après avoir importé le \n"
"        fichier et rempli le formulaire"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
