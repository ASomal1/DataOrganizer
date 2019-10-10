# -*- coding: utf-8 -*-
"""

@author: as4121
"""

# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'gg.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QMessageBox, qApp, QMenu, QAction
import pandas as pd
import numpy as np
import pickle
import os


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1022, 812)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        MainWindow.setFont(font)
        MainWindow.setMouseTracking(False)
        MainWindow.setAutoFillBackground(False)
        MainWindow.setStyleSheet("background-color: rgb(44, 44, 44);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setAutoFillBackground(False)
        self.centralwidget.setStyleSheet(".bg_color{ background-color: #334422; }")
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(630, 90, 201, 41))
        self.pushButton.clicked.connect(self.uploadxlsfile)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        self.pushButton.setStyleSheet("background-color: rgb(48, 150, 175);\n"
"color: rgb(255, 255, 255);")
        self.pushButton.setAutoDefault(False)
        self.pushButton.setDefault(False)
        self.pushButton.setFlat(False)
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(300, 710, 201, 41))
        self.pushButton_2.setEnabled(False)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(12) 
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        self.pushButton_2.setAutoFillBackground(False)
        self.pushButton_2.setStyleSheet("background-color: rgb(48, 150, 175);\n"
"color: rgb(255, 255, 255);")
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.convert)
        
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(550, 710, 201, 41))
        self.pushButton_3.setEnabled(False)
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(12) 
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        self.pushButton_3.setAutoFillBackground(False)
        self.pushButton_3.setStyleSheet("background-color: rgb(48, 150, 175);\n"
"color: rgb(255, 255, 255);")
        self.pushButton_3.setObjectName("pushButton_2")
        self.pushButton_3.clicked.connect(self.showAllColumns)
        
        self.tableWidget = QtWidgets.QTableView(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(80, 180, 841, 511))
        self.tableWidget.setAutoFillBackground(True)
        self.tableWidget.setStyleSheet("background-color: rgb(225, 225, 225);")
        self.tableWidget.setObjectName("tableWidget")
       
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(90, 40, 311, 31))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setStyleSheet("color: rgb(255, 255, 255);")
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(160, 80, 421, 31))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setStyleSheet("color: rgb(191, 255, 191);\n"
"")
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(160, 130, 421, 31))
        font = QtGui.QFont()
        font.setFamily("Trebuchet MS")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_3.setFont(font)
        self.label_3.setStyleSheet("color: rgb(191, 255, 191);\n"
"")
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(110, 80, 31, 31))
        self.label_4.setStyleSheet("background-image: url(:/newPrefix/excel-3-32.png);")
        self.label_4.setText("")
        
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(110, 130, 31, 31))
        self.label_5.setStyleSheet("background-image: url(:/newPrefix/csv-32.png);")
        self.label_5.setText("")
        self.label_5.setObjectName("label_5")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1022, 31))
        self.label_4.hide()
        self.label_5.hide()
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.menubar.setFont(font)
        self.menubar.setAutoFillBackground(False)
        self.menubar.setStyleSheet("color: rgb(0, 0, 0); background-color: rgb(48, 150, 175);")
        self.menubar.setObjectName("menubar")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setAutoFillBackground(False)
        self.menuHelp.setStyleSheet("color: rgb(255, 255, 255);")
        self.menuHelp.setObjectName("menuHelp")
        self.menuHelp.addAction('Get Help').triggered.connect(self.helpDialogBox)
        self.menuExit = QtWidgets.QMenu(self.menubar)
        self.menuExit.setObjectName("menuExit")
        self.menuExit.addAction('Exit').triggered.connect(qApp.closeAllWindows)
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionGet_Help = QtWidgets.QAction(MainWindow)
        self.actionGet_Help.setObjectName("actionGet_Help")
        self.menubar.addAction(self.menuHelp.menuAction())
        self.menubar.addAction(self.menuExit.menuAction())
        self.tableWidget.doubleClicked.connect(self.doubleClickedWork)
        self.retranslateUi(MainWindow)
        self.removinglist=[]
        
        self.tableWidget.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        
        self.popup_menu = QMenu(self.tableWidget)
        self.hideColumn = QAction("Hide Column", self.tableWidget)
        self.hideColumn.triggered.connect(self.hideTableColumn)
        self.popup_menu.addAction(self.hideColumn)
        
        
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
    
    def hideTableColumn(self):
         indexes = self.tableWidget.selectedIndexes()
         self.colindex=0;
         for index in indexes:
            self.colindex=index.column()
         if self.colindex != 0:
            self.tableWidget.hideColumn(self.colindex)
            self.removinglist.append(self.colindex)

        
         self.colhead = self.pivotdata.head(self.pivotdata.shape[1])
         self.colhead.index
         
         
    
    def _ctx_menu_cb(self, pos):
        #print ("pos: "+str( pos))
        self.popup_menu.exec_(self.tableWidget.mapToGlobal(pos))  
        
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Data Organizer"))
        self.pushButton.setToolTip("<i><font color='#2D96AF'>Please select an excel and csv file</font></i>")
        self.pushButton_2.setToolTip("<i><font color='#2D96AF'>This will convert the data and output a pivot table. To download the file, click on the first column of the pivot table</font></i>")
                                     
        self.pushButton.setText(_translate("MainWindow", "Upload Files"))
        self.pushButton_2.setText(_translate("MainWindow", "Convert Data"))
        self.pushButton_3.setText(_translate("MainWindow", "Show All Columns"))
        self.pushButton_3.setToolTip("<i><font color='#2D96AF'>Please click here to show all the available columns in the table</font></i>")
        self.label.setText(_translate("MainWindow", "Selected Files:"))
        
        self.menuHelp.setTitle(_translate("MainWindow", "Help"))
        self.menuExit.setTitle(_translate("MainWindow", "Exit"))
        self.actionGet_Help.setText(_translate("MainWindow", "Get Help"))
    
    def doubleClickedWork(self):
        self.tempdata=self.tempdata.drop(self.tempdata.columns[self.removinglist],axis=1)
        indexes = self.tableWidget.selectedIndexes()
        for index in indexes:
            self.colindex=index.column()
        if self.colindex == 0:
            buttonReply = QMessageBox.question(None, 'Confirm download', "Do you want to download the calculated data?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if buttonReply == QMessageBox.Yes:
                #print('Yes clicked.')
                self.saveFileDialog()
     
        
    def showAllColumns(self):  
            
            self.model = PandasModel(self.pivotdata)
            self.tableWidget.setModel(self.model)
            self.removinglist.clear()
            #print(self.model.columnCount())
            
            for i in range(0,self.model.columnCount()):
                 self.tableWidget.showColumn(i)
        
    
    def convert(self):
            self.df= pd.concat([self.dfexcel,self.dfcsv], sort=False)
            
            #create pivot of the data
            self.pivotdata= pd.pivot_table(self.df, index= ["Client ID"],
                                           values=["Book","Buy/Sell","Name","Instrument Code","Shares","Gross Price",
                                                   "Settlement Date","Ticker","State","Gross Value","Commission",
                                                   "Settlement Month" ,"Unnamed: 13","Unnamed: 14","Unnamed: 15",
                                                   "Unnamed: 16","Unnamed: 17","Search For","Length"],
                                           aggfunc={"Book":len,"Buy/Sell":len,"Name":len,"Instrument Code":len,
                                                    "Shares":np.sum,"Gross Price":np.sum,"Settlement Date":len,
                                                    "Ticker":len,"State":len,"Gross Value":np.sum,"Commission":len,
                                                  "Settlement Month":len ,"Unnamed: 13":len,"Unnamed: 14":len,
                                                    "Unnamed: 15":len,"Unnamed: 16":len,"Unnamed: 17":len,"Search For":len,"Length":len},
                                           fill_value=0
                                           )
            
         
            self.head = self.pivotdata.head(self.pivotdata.shape[0])
            self.cl_id = []
            for row in self.head.index: 
                self.cl_id.append(row) 
            #print(self.cl_id)
            self.pivotdata.insert(0, "Client ID", self.cl_id, True) 
            self.pickle_out = open("data.pickle","wb")
            pickle.dump(self.pivotdata, self.pickle_out)
            self.pickle_out.close()
            
            self.pickle_in = open("data.pickle","rb")
            self.example_dict = pickle.load(self.pickle_in)
            
            
           
            self.model = PandasModel(self.pivotdata)
            self.tempdata=self.pivotdata
            self.tableWidget.setModel(self.model)  
            self.tableWidget.customContextMenuRequested.connect(self._ctx_menu_cb)
            self.pushButton_3.setEnabled(True)
            
    def saveFileDialog(self):
        options = QFileDialog.Options()
        
        fileName, _ = QFileDialog.getSaveFileName( self.centralwidget,"QFileDialog.getSaveFileName()","","csv Files (*.csv)", options=options)
        if fileName:
            export_csv = self.tempdata.to_csv (r''+fileName, index = None, header=True)
            QMessageBox.information(None, "Download status",
                " File "+fileName+" has been downloaded successfully")
            
    def helpDialogBox(self):
        QMessageBox.information(None, "Help",
                "1) Click upload button to select two files: one is excel fiel and another is csv file." + "\n" +
                "2) Click convert button to merge the selected files and display the data." + "\n"+
                "3) After converting data, Right Click on the cell in the table to remove the cell's column." + "\n"+
                "4) Click on 'Show all columns' button to show all the available data columns in table." + "\n"+
                "5) The merged data can be downloaded by double clicking on the first column of the table shown.")
            
    
    def uploadxlsfile(self):
        
        
        self.file = QFileDialog.getOpenFileNames( self.centralwidget,"Select csv and excel files(maximum 2 files of different type)", "/",
                                                "Excel files (*.xls *.xlsx *.csv *.XLS *.XLSX *.CSV)")
        self.lst=self.file[0]
        
        #print(self.lst)
        #print(len(self.lst))
        if(len(self.lst) <2):
           
            QMessageBox.critical(None, "Invalid File selection",
                " select atleast 2 files to proceed")
            
           
            return
        if(len(self.lst) >2):
           
            QMessageBox.critical(None, "Invalid File selection",
                "only 2 files can be selected")
            return
        
        
        
            
        
        if((self.lst[0].endswith('xls') or self.lst[0].endswith('xlsx')) and self.lst[1].endswith('csv') ):
            
            self.dfexcel = pd.read_excel (self.lst[0]) 
            self.dfcsv = pd.read_csv(self.lst[1])
            headcsv, tailcsv = os.path.split(self.lst[1])
            headexcel, tailexcel = os.path.split(self.lst[0])
            self.label_4.show()
            self.label_5.show()
            self.label_2.setText(tailexcel)
            self.label_3.setText(tailcsv)
            self.pushButton.setEnabled(False)
            self.pushButton_2.setEnabled(True)
           
            
        elif ((self.lst[1].endswith('xls') or self.lst[1].endswith('xlsx')) and self.lst[0].endswith('csv') ):    
            self.dfexcel = pd.read_excel (self.lst[1]) 
            self.dfcsv = pd.read_csv(self.lst[0])
            headcsv, tailcsv = os.path.split(self.lst[0])
            headexcel, tailexcel = os.path.split(self.lst[1])
            self.label_4.show()
            self.label_5.show()
            self.label_2.setText(tailexcel)
            self.label_3.setText(tailcsv)
            self.pushButton.setEnabled(False)
            self.pushButton_2.setEnabled(True)
            
            
        else:
             QMessageBox.critical(None, "File type error",
                "only excel formatted (xls, xlsx, csv) files are allowed to upload")
    

class PandasModel(QtCore.QAbstractTableModel):
    """
    Class to populate a table view with a pandas dataframe
    """
    def __init__(self, data, parent=None):
        QtCore.QAbstractTableModel.__init__(self, parent)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if index.isValid():
            if role == QtCore.Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == QtCore.Qt.Horizontal and role == QtCore.Qt.DisplayRole:
            return self._data.columns[col]
        return None            

import resource_rc


import sys



app = QtWidgets.QApplication(sys.argv)
ex = Ui_MainWindow()
w = QtWidgets.QMainWindow()
ex.setupUi(w)
w.show()
sys.exit(app.exec_())