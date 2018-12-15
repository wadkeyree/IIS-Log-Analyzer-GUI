# 12/15/2018 
# Developer: Elsadig Mohamed
# This is a Simple GUI Application built on top of the Microsoft LogParser 2.2 to provide an easiser way to write queries and pull data from IIS logs.

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
import sys, os, csv
from subprocess import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import pandas as pd
import time
import shutil


class Ui_MainWindow(QWidget):
    def setupUi(self, MainWindow):

        # Defining global veriables
        self.path=''
        self.filename=''  
        self.FILE=''     
        self.GIF='' 
        self.CSV=''
        self.COUNT = ''
        self.NEW_Working_DIR = '' 
        self.item_selected=''  
        self.scripts = []
        self.return_code=None 
        
        # Window widgets and objects were generated using the QtPy5 Designer.
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1197, 890)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setEnabled(False)
        font = QtGui.QFont()
        
        font.setPointSize(10)
        self.lineEdit.setFont(font)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout.addWidget(self.lineEdit)
        self.btn_browse = QtWidgets.QPushButton(self.centralwidget)
        self.btn_browse.setEnabled(False)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.btn_browse.setFont(font)
        self.btn_browse.setObjectName("btn_browse")
        self.horizontalLayout.addWidget(self.btn_browse)
        self.cbSelectQuery = QtWidgets.QComboBox(self.centralwidget)
        self.cbSelectQuery.setEnabled(False)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.cbSelectQuery.setFont(font)
        self.cbSelectQuery.setObjectName("cbSelectQuery")
        self.cbSelectQuery.addItem("Select Custom Query")
        self.cbSelectQuery.addItem("Generate CSV")
        self.cbSelectQuery.addItem("Pie3D Chart")
        self.cbSelectQuery.addItem("Pie Chart")
        self.cbSelectQuery.addItem("Line Chart")
        self.cbSelectQuery.addItem("Column3D Chart")
        self.cbSelectQuery.addItem("Col-Clust Chart")
        #
        # self.cbSelectQuery.model().item(5).setEnabled(False)
        self.cbSelectQuery.addItem("")
        self.horizontalLayout.addWidget(self.cbSelectQuery)
        self.btnClear = QtWidgets.QPushButton(self.centralwidget)
        self.btnClear.setEnabled(False)
        self.btnClear.setFont(font)
        self.btnClear.setObjectName("btnClear")
        self.horizontalLayout.addWidget(self.btnClear)
        self.btnRun = QtWidgets.QPushButton(self.centralwidget)
        self.btnRun.setEnabled(False)
        Tfont = QtGui.QFont()

        Tfont.setFamily('Courier')
        Tfont.setFixedPitch(True)
        Tfont.setPointSize(10)

        font.setBold(False)
        font.setWeight(50)
        self.btnRun.setFont(font)
        self.btnRun.setObjectName("btnRun")
        self.horizontalLayout.addWidget(self.btnRun)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.splitter_2 = QtWidgets.QSplitter(self.centralwidget)
        self.splitter_2.setOrientation(QtCore.Qt.Vertical)
        self.splitter_2.setObjectName("splitter_2")
        self.splitter = QtWidgets.QSplitter(self.splitter_2)
        self.splitter.setOrientation(QtCore.Qt.Horizontal)
        self.splitter.setObjectName("splitter")
        self.textEdit = QtWidgets.QTextEdit(self.splitter)
        self.textEdit.setFont(Tfont)
        # Adding syntax highlighting 
        self.highlighter = Highlighter(self.textEdit.document())
        self.textEdit.setEnabled(False)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        font.setStyleStrategy(QtGui.QFont.PreferAntialias)
        
        self.textEdit.setObjectName("textEdit")

        self.graphicsView = QtWidgets.QGraphicsView(self.splitter)

        self.graphicsView.setObjectName("graphicsView")


        self.tableView = QtWidgets.QTableView(self.splitter_2)
        self.tableView.setObjectName("tableView")
        self.chb_viewExcel = QtWidgets.QCheckBox(self.splitter_2)
        self.chb_viewExcel.setEnabled(False)
        self.chb_viewExcel.setObjectName("chb_viewExcel")
        self.verticalLayout.addWidget(self.splitter_2)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1197, 21))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionOpen = QtWidgets.QAction(MainWindow)
        self.actionOpen.setStatusTip("")
        self.actionOpen.setObjectName("actionOpen")
        self.actionRun_Custom = QtWidgets.QAction(MainWindow)
        self.actionRun_Custom.setStatusTip("")
        self.actionRun_Custom.setVisible(False)
        self.actionRun_Custom.setObjectName("actionRun_Custom")
        self.actionCreate_Query = QtWidgets.QAction(MainWindow)
        self.actionCreate_Query.setStatusTip("")
        self.actionCreate_Query.setVisible(False)
        self.actionCreate_Query.setObjectName("actionCreate_Query")
        self.actionExit = QtWidgets.QAction(MainWindow)
        self.actionExit.setStatusTip("")
        self.actionExit.setObjectName("actionExit")
        self.actionAbout = QtWidgets.QAction(MainWindow)
        self.actionAbout.setStatusTip("")
        self.actionAbout.setObjectName("actionAbout")
        self.menuFile.addAction(self.actionOpen)
        self.menuFile.addAction(self.actionRun_Custom)
        self.menuFile.addAction(self.actionCreate_Query)
        self.menuFile.addAction(self.actionExit)
        self.menuHelp.addAction(self.actionAbout)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

       
        #Generating SIGNAL to connection to functions.
  
        self.actionAbout.triggered.connect(self.about)
        self.actionExit.triggered.connect(self.exit)
        self.actionOpen.triggered.connect(self.open_file)
        self.btn_browse.clicked.connect(self.browse)
        self.actionRun_Custom.triggered.connect(self.enable_funcs)
        self.chb_viewExcel.stateChanged.connect(self.save_script)
        self.btnRun.clicked.connect(self.run)
        self.btnClear.clicked.connect(self.clear)
        self.cbSelectQuery.currentIndexChanged.connect(self.run_custom_query)
    # The clear method to clear the editor text.
    def clear(self):
        self.textEdit.clear()

    def loadCsv(self, _csv):
        try: 
            self.lineEdit.setText('{}'.format(_csv))
            df = pd.read_csv(_csv)
            model = PandasModel(df)
            self.tableView.setModel(model) 
        except FileNotFoundError as e:
            pass   
    # This method runs the main logic in run_custom_code.
    def run(self):
        self.run_custom_query()       
        if self.return_code == 0 or self.return_code == 1:
            font = QtGui.QFont()
            font.setPointSize(12)
            font.setBold
            self.chb_viewExcel.setFont(font)
            self.chb_viewExcel.setEnabled(True)
            self.chb_viewExcel.setChecked(False)
        else:
            self.chb_viewExcel.setEnabled(False)
            font = QtGui.QFont()
            font.setPointSize(10)
    # This method is responsible for formulating the queries.  
    def custom_queries (self, selection):
        
        
        if selection == 'Pie Chart' or selection =='Pie3D Chart' or selection =='Line Chart' or selection =='Colmun3D Chart' or selection =='Col-Clust Chart':
            self.item_selected = selection
        else:
            pass
        
        if selection == 'Generate CSV': 
           
            if self.return_code == None:            
                return  ' SELECT * INTO outPut.csv FROM IIS.log WHERE cs-uri-stem LIKE '"'%cta%'"' '
            else:
                return  self.textEdit.toPlainText().replace('outPut.gif', 'outPut.csv')                

        elif selection == 'Pie3D Chart' or selection == 'Pie Chart' or selection == 'Line Chart'  or selection == 'Column3D Chart' or selection == 'Col-Clust Chart': 
            if 'outPut.csv' in self.textEdit.toPlainText() or self.return_code == None:
                return ' SELECT cs-uri-stem, MAX(time-taken) INTO outPut.gif FROM IIS.log GROUP BY cs-uri-stem ORDER BY MAX(time-taken) '
            else:
                return self.textEdit.toPlainText().replace('outPut.csv', 'outPut.gif')
        elif self.textEdit.toPlainText(None):
            pass
    #This method implements the logic for enabling/ disabling objects such buttons etc.
    def enable_funcs(self):
        
        #self.actionCreate_Query.setVisible(True)
        self.cbSelectQuery.setEnabled(True)
        self.btn_browse.setEnabled(True)
        self.btnClear.setEnabled(True)
        self.lineEdit.setEnabled(True)
      
        self.textEdit.setEnabled(True)

    def find_file(self, path):
        for root, dirs, files in os.walk(path):
            if name in files:
                return os.path.join(root, name)  
    # Here we attemp to save a query although some more work is needed here.
    def save_script(self, scripts):       
        
        if len(self.scripts) >=1:
            
            scriptsFile = os.path.join(os.path.join(self.NEW_Working_DIR, os.pardir), 'scriptsFile.txt')
            with open(scriptsFile, 'w') as input_f:
                for script in self.scripts:
                    input_f.write(script)
        if self.chb_viewExcel.stateChanged: 
                       
            font = QtGui.QFont()
            font.setPointSize(10)            
            self.chb_viewExcel.setFont(font)
            self.chb_viewExcel.setText('Query has been saved successfully')
    # Here the main logic for running the query is implemented.
    def run_custom_query(self): 
        self.cbSelectQuery.model().item(0).setEnabled(False)

        logs = self.NEW_Working_DIR + '\*.log'
        
        render_csv = ' " -i:W3C " '
        render_gif = {'pie':' " -o:chart -chartType:PieExploded -categories:off ', 'pie3D': ' " -o:chart -chartType:PieExploded3D -categories:off',  \
        'line':' " -o:chart -chartType:Line ', 'Col3D': ' " -o:chart -chartType:Column3D ', 'colClust': ' " -o:chart -chartType:ColumnClustered '}            
        script = self.custom_queries(self.cbSelectQuery.currentText())
        
        try:
     
            if self.cbSelectQuery.currentText() == "Generate CSV":
                cmd = 'logparser.exe ' '"' + script.replace('outPut.csv', self.CSV).replace('IIS.log', str(logs)) + '"'  " -i:W3C "
                
            elif  self.cbSelectQuery.currentText() == "Col-Clust Chart":  
                cmd = 'logparser.exe  -i IISW3C ' '"' + script.replace('outPut.gif', self.GIF).replace('IIS.log', str(logs)) + render_gif['colClust'] 

            elif  self.cbSelectQuery.currentText() == "Column3D Chart":  
                cmd = 'logparser.exe  -i IISW3C ' '"' + script.replace('outPut.gif', self.GIF).replace('IIS.log', str(logs)) + render_gif['Col3D'] 
        

            elif  self.cbSelectQuery.currentText() == "Line Chart":  
                cmd = 'logparser.exe  -i IISW3C ' '"' + script.replace('outPut.gif', self.GIF).replace('IIS.log', str(logs)) + render_gif['line']             

            elif self.cbSelectQuery.currentText() == "Pie Chart":
                cmd = 'logparser.exe  -i IISW3C ' '"' + script.replace('outPut.gif', self.GIF).replace('IIS.log', str(logs)) + render_gif['pie']

            elif self.cbSelectQuery.currentText() == "Pie3D Chart":
                cmd = 'logparser.exe  -i IISW3C ' '"' + script.replace('outPut.gif', self.GIF).replace('IIS.log', str(logs)) + render_gif['pie3D']
            else:
                pass
        except:
            pass
        
        
        if self.cbSelectQuery.currentText() == "Pie3D Chart" or self.cbSelectQuery.currentText() == "Generate CSV"  or self.cbSelectQuery.currentText() == "Column3D Chart"  \
            or self.cbSelectQuery.currentText() == "Pie Chart" or self.cbSelectQuery.currentText() == "Line Chart" or self.cbSelectQuery.currentText() == "Col-Clust Chart":
            self.btnRun.setEnabled(True)
            self.btnClear.setEnabled(True)
        else:
            self.btnRun.setEnabled(False)
            self.btnClear.setEnabled(False)
        
        try:  
            print(cmd)                   
            with Popen(cmd, cwd=self.NEW_Working_DIR, stdin=PIPE, stdout=PIPE, stderr=STDOUT, shell=True) as P:  
                
                               
                res = P.communicate() 
                if P.returncode == 0 or P.returncode == 1:  
                    self.return_code = P.returncode                            
                    self.chb_viewExcel.setText('Check box to save your query ! [] Query executed successfully, returned Code: {}'.format(P.returncode))  

                    #Assign the script to the global variable self.scripts
                    self.scripts.append(script)                    
                    self.textEdit.setText('{}'.format(script))
                   
                    self.tableView.setModel(None)
                    self.graphicsView.setScene(None)
                    #self.cbSelectQuery.model().item(5).setEnabled(True)                
                    if self.CSV in cmd:
                        self.loadCsv(self.CSV)
                    elif self.GIF in cmd:
                        self.display_image(self.GIF)
                else:
                    self.status_msg(str(res[0].decode(encoding='utf-8')).replace('\\n',''))
                    self.chb_viewExcel.setText('Check box to save your query ! [] Query execution failed, returned Code: {}'.format(P.returncode))
        except UnboundLocalError as e:
            self.status_msg('{}'.format(e))
    #Here we run the Open_file function.
    def browse(self):
        self.open_file()
    # This method is where the Chart logic dispaly is imlemented.
    def display_image(self, _gif):
        try:
            scene =  QGraphicsScene()
            scene.addPixmap(QPixmap(_gif))
            self.graphicsView.setScene(scene)
            self.graphicsView.show() 
            self.lineEdit.setText('{}'.format(_gif))
        except:
            self.graphicsView.setScene(None)
 
    # This method is where we get the open file dialog to open when its selected in the Menu or browse button.
    def open_file(self):
        fileNames, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Open File", "C:\\Users", "Log Files (*.log)")
     
        bad_files = []
        good_files = [] 
        for count, file_name in enumerate(fileNames):    
            with open(file_name, 'r') as myfile:
                heads = [next(myfile) for x in range(5)] 
                    
                for head in heads:      
                    if head.startswith("#Fields:"):                                         
                        good_files.append(file_name)
                        _time = time.strftime("%Y%m%d-%H%M%S")
                    else:
                        bad_files.append(file_name)
        if good_files:
            dir_name, file_name = os.path.split(os.path.abspath(fileNames[0])) 
                                       
            self.NEW_Working_DIR =  dir_name #os.path.join(dir_name, _time)
            print(self.NEW_Working_DIR)
            '''try:
                os.mkdir(self.NEW_Working_DIR)                                
            except TypeError as e:
                self.status_msg(e)
                
            for good_file in good_files:
                try:
                    shutil.move(os.path.abspath(good_file), self.NEW_Working_DIR)
                except TypeError as e:
                    self.status_msg(e)'''

                     
            self.FILE = os.path.join(self.NEW_Working_DIR, file_name)              
            self.GIF = self.FILE.replace('.log','.gif').replace(' ','') 
            self.CSV = self.FILE.replace('.log','.csv').replace(' ','')     
            self.status_msg('{} File(s) selected for Analysis \n{}Good  Files \n{} Bad Files (A bad file is one missing a "#Fields" header) \n \n \n Please select \
an option from the drop down list'. format((count+1), len(good_files), (count+1) - len(good_files)))                                                  
            self.lineEdit.setText(self.NEW_Working_DIR)
            self.enable_funcs()
            self.btnRun.setEnabled(False)
            self.COUNT = count
  
    # This is for the exit functionality found  under the main menu
    def exit(self):
        reply = QMessageBox.question(self, "Message","Are you sure you want to quit? Any unsaved work will be lost.",
            QMessageBox.Save | QMessageBox.Close | QMessageBox.Cancel,
            QMessageBox.Save)

        if reply == QMessageBox.Close:
            sys.exit()
        else:
            pass
    # A method that dispalays a pop-up message with whatever text passed to it and its used in a few places through the app.
    def status_msg(self, in_msg):
        
        msg = QMessageBox()
        msg.setText(in_msg)
        msg.exec_()

    # For the About item under the menu.
    def about(self):
        
        msg = QMessageBox()
        msg.setText("Developer: Elsadig Mohamed. \n \n This utility provide away to query IIS logs")
        msg.exec_()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Trubo Logger for SDL Support by Elsadig Mohamed"))
        self.btn_browse.setText(_translate("MainWindow", "Browse"))
        self.cbSelectQuery.setToolTip(_translate("MainWindow", "Select the custom query to run"))
        self.cbSelectQuery.setItemText(0, _translate("MainWindow", "Select Custom Query"))        
        self.cbSelectQuery.setItemText(1, _translate("MainWindow", "Generate CSV"))
        self.cbSelectQuery.setItemText(2, _translate("MainWindow", "Pie3D Chart"))
        self.cbSelectQuery.setItemText(3, _translate("MainWindow", "Pie Chart"))
        self.cbSelectQuery.setItemText(4, _translate("MainWindow", "Line Chart"))
        self.cbSelectQuery.setItemText(5, _translate("MainWindow", "Column3D Chart"))
        self.cbSelectQuery.setItemText(6, _translate("MainWindow", "Col-Clust Chart"))
        self.btnClear.setText(_translate("MainWindow", "Clear"))
        self.btnRun.setToolTip(_translate("MainWindow", "Run query"))
        self.btnRun.setText(_translate("MainWindow", "Run"))
        self.chb_viewExcel.setText(_translate("MainWindow", "View Excel if available"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.menuHelp.setTitle(_translate("MainWindow", "Help"))
        self.actionOpen.setText(_translate("MainWindow", "Open"))
        self.actionOpen.setToolTip(_translate("MainWindow", "Open IIS log to be analyzed."))
        self.actionRun_Custom.setText(_translate("MainWindow", "Run Custom"))
        self.actionRun_Custom.setToolTip(_translate("MainWindow", "Run a custom query from the drop down list."))
        self.actionCreate_Query.setText(_translate("MainWindow", "Create Query"))
        self.actionCreate_Query.setToolTip(_translate("MainWindow", "Create an run your own query."))
        self.actionExit.setText(_translate("MainWindow", "Exit"))
        self.actionAbout.setText(_translate("MainWindow", "About"))

''' This class will handle the syntax formatting logic for SQL more work can be done here as well.'''

class Highlighter(QSyntaxHighlighter):
    def __init__(self, parent=None):
        super(Highlighter, self).__init__(parent)

        keywordFormat = QTextCharFormat()
        keywordFormat.setForeground(Qt.blue)
        keywordFormat.setFontWeight(QFont.Bold)

        keywordPatterns = ["\\badd\\b","\\bEXTERNAL\\b","\\bPROCEDURE\\b","\\bALL\\b","\\bFETCH\\b","\\bPUBLIC\\b","\\bALTER\\b","\\bFILE\\b","\\bRAISERROR\\b","\\bAND\\b", \
        "\\bFILLFACTOR\\b","\\bREAD\\b","\\bANY\\b","\\bFOR\\b","\\bREADTEXT\\b","\\bAS\\b","\\bFOREIGN\\b","\\bRECONFIGURE\\b","\\bASC\\b","\\bFREETEXT\\b","\\bREFERENCES\\b", \
        "\\bAUTHORIZATION\\b","\\bFREETEXTTABLE\\b","\\bREPLICATION\\b","\\bBACKUP\\b","\\bFROM\\b","\\bRESTORE\\b","\\bBEGIN\\b","\\bFULL\\b","\\bRESTRICT\\b", "\\bBETWEEN\\b","\\bFUNCTION\\b", \
        "\\bRETURN\\b","\\bBREAK\\b","\\bGOTO\\b","\\bREVERT\\b","\\bBROWSE\\b","\\bGRANT\\b","\\bREVOKE\\b","\\bBULK\\b","\\bGROUP\\b","\\bRIGHT\\b","\\bBY\\b","\\bHAVING\\b","\\bROLLBACK\\b", \
        "\\bCASCADE\\b","\\bHOLDLOCK\\b","\\bROWCOUNT\\b","\\bCASE\\b","\\bIDENTITY\\b","\\bROWGUIDCOL\\b","\\bCHECK\\b","\\bIDENTITY_INSERT\\b","\\bRULE\\b","\\bCHECKPOINT\\b","\\bIDENTITYCOL\\b", \
        "\\bSAVE\\b","\\bCLOSE\\b","\\bIF\\b","\\bSCHEMA\\b","\\bCLUSTERED\\b","\\bIN\\b","\\bSECURITYAUDIT\\b","\\bCOALESCE\\b","\\bINDEX\\b","\\bSELECT\\b","\\bCOLLATE\\b","\\bINNER\\b", \
        "\\bSEMANTICKEYPHRASETABLE\\b","\\bCOLUMN\\b","\\bINSERT\\b","\\bSEMANTICSIMILARITYDETAILSTABLE\\b","\\bCOMMIT\\b","\\bINTERSECT\\b","\\bSEMANTICSIMILARITYTABLE\\b","\\bCOMPUTE\\b","\\bINTO\\b", \
        "\\bSESSION_USER\\b","\\bCONSTRAINT\\b","\\bIS\\b","\\bSET\\b","\\bCONTAINS\\b","\\bJOIN\\b","\\bSETUSER\\b","\\bCONTAINSTABLE\\b","\\bKEY\\b","\\bSHUTDOWN\\b","\\bCONTINUE\\b","\\bKILL\\b", \
        "\\bSOME\\b","\\bCONVERT\\b","\\bLEFT\\b","\\bSTATISTICS\\b","\\bCREATE\\b","\\bLIKE\\b","\\bSYSTEM_USER\\b","\\bCROSS\\b","\\bLINENO\\b","\\bTABLE\\b","\\bCURRENT\\b","\\bLOAD\\b","\\bTABLESAMPLE\\b", \
        "\\bCURRENT_DATE\\b","\\bMERGE\\b","\\bTEXTSIZE\\b","\\bCURRENT_TIME\\b","\\bNATIONAL\\b","\\bTHEN\\b","\\bCURRENT_TIMESTAMP\\b","\\bNOCHECK\\b","\\bTO\\b","\\bCURRENT_USER\\b","\\bNONCLUSTERED\\b", \
        "\\bTOP\\b","\\bCURSOR\\b","\\bNOT\\b","\\bTRAN\\b","\\bDATABASE\\b","\\bNULL\\b","\\bTRANSACTION\\b","\\bDBCC\\b","\\bNULLIF\\b","\\bTRIGGER\\b","\\bDEALLOCATE\\b","\\bOF\\b","\\bTRUNCATE\\b", \
        "\\bDECLARE\\b","\\bOFF\\b","\\bTRY_CONVERT\\b","\\bDEFAULT\\b","\\bOFFSETS\\b","\\bTSEQUAL\\b","\\bDELETE\\b","\\bON\\b","\\bUNION\\b","\\bDENY\\b","\\bOPEN\\b","\\bUNIQUE\\b","\\bDESC\\b", \
        "\\bOPENDATASOURCE\\b","\\bUNPIVOT\\b","\\bDISK\\b","\\bOPENQUERY\\b","\\bUPDATE\\b","\\bDISTINCT\\b","\\bOPENROWSET\\b","\\bUPDATETEXT\\b","\\bDISTRIBUTED\\b","\\bOPENXML\\b","\\bUSE\\b", \
        "\\bDOUBLE\\b","\\bOPTION\\b","\\bUSER\\b","\\bDROP\\b","\\bOR\\b","\\bVALUES\\b","\\bDUMP\\b","\\bORDER\\b","\\bVARYING\\b","\\bELSE\\b","\\bOUTER\\b","\\bVIEW\\b","\\bEND\\b","\\bOVER\\b", \
        "\\bWAITFOR\\b","\\bERRLVL\\b","\\bPERCENT\\b","\\bWHEN\\b","\\bESCAPE\\b","\\bPIVOT\\b","\\bWHERE\\b","\\bEXCEPT\\b","\\bPLAN\\b","\\bWHILE\\b","\\bEXEC\\b","\\bPRECISION\\b","\\bWITH\\b", \
        "\\bEXECUTE\\b","\\bPRIMARY\\b","\\bWITHINGROUP\\b","\\bEXISTS\\b","\\bPRINT\\b","\\bWRITETEXT\\b","\\bEXIT\\b","\\bPROC\\b" \
        "\\badd\\b","\\bexternal\\b","\\bprocedure\\b","\\ball\\b","\\bfetch\\b","\\bpublic\\b","\\balter\\b","\\bfile\\b","\\braiserror\\b","\\band\\b", \
        "\\bfillfactor\\b","\\bread\\b","\\bany\\b","\\bfor\\b","\\breadtext\\b","\\bas\\b","\\bforeign\\b","\\breconfigure\\b","\\basc\\b","\\bfreetext\\b","\\breferences\\b", \
        "\\bauthorization\\b","\\bfreetexttable\\b","\\breplication\\b","\\bbackup\\b","\\bfrom\\b","\\brestore\\b","\\bbegin\\b","\\bfull\\b","\\brestrict\\b", "\\bbetween\\b","\\bfunction\\b", \
        "\\breturn\\b","\\bbreak\\b","\\bgoto\\b","\\brevert\\b","\\bbrowse\\b","\\bgrant\\b","\\brevoke\\b","\\bbulk\\b","\\bgroup\\b","\\bright\\b","\\bby\\b","\\bhaving\\b","\\brollback\\b", \
        "\\bcascade\\b","\\bholdlock\\b","\\browcount\\b","\\bcase\\b","\\bidentity\\b","\\browguidcol\\b","\\bcheck\\b","\\bidentity_insert\\b","\\brule\\b","\\bcheckpoint\\b","\\bidentitycol\\b", \
        "\\bsave\\b","\\bclose\\b","\\bif\\b","\\bschema\\b","\\bclustered\\b","\\bin\\b","\\bsecurityaudit\\b","\\bcoalesce\\b","\\bindex\\b","\\bselect\\b","\\bcollate\\b","\\binner\\b", \
        "\\bsemantickeyphrasetable\\b","\\bcolumn\\b","\\binsert\\b","\\bsemanticsimilaritydetailstable\\b","\\bcommit\\b","\\bintersect\\b","\\bsemanticsimilaritytable\\b","\\bcompute\\b","\\binto\\b", \
        "\\bsession_user\\b","\\bconstraint\\b","\\bis\\b","\\bset\\b","\\bcontains\\b","\\bjoin\\b","\\bsetuser\\b","\\bcontainstable\\b","\\bkey\\b","\\bshutdown\\b","\\bcontinue\\b","\\bkill\\b", \
        "\\bsome\\b","\\bconvert\\b","\\bleft\\b","\\bstatistics\\b","\\bcreate\\b","\\blike\\b","\\bsystem_user\\b","\\bcross\\b","\\blineno\\b","\\btable\\b","\\bcurrent\\b","\\bload\\b","\\btablesample\\b", \
        "\\bcurrent_date\\b","\\bmerge\\b","\\btextsize\\b","\\bcurrent_time\\b","\\bnational\\b","\\bthen\\b","\\bcurrent_timestamp\\b","\\bnocheck\\b","\\bto\\b","\\bcurrent_user\\b","\\bnonclustered\\b", \
        "\\btop\\b","\\bcursor\\b","\\bnot\\b","\\btran\\b","\\bdatabase\\b","\\bnull\\b","\\btransaction\\b","\\bdbcc\\b","\\bnullif\\b","\\btrigger\\b","\\bdeallocate\\b","\\bof\\b","\\btruncate\\b", \
        "\\bdeclare\\b","\\boff\\b","\\btry_convert\\b","\\bdefault\\b","\\boffsets\\b","\\btsequal\\b","\\bdelete\\b","\\bon\\b","\\bunion\\b","\\bdeny\\b","\\bopen\\b","\\bunique\\b","\\bdesc\\b", \
        "\\bopendatasource\\b","\\bunpivot\\b","\\bdisk\\b","\\bopenquery\\b","\\bupdate\\b","\\bdistinct\\b","\\bopenrowset\\b","\\bupdatetext\\b","\\bdistributed\\b","\\bopenxml\\b","\\buse\\b", \
        "\\bdouble\\b","\\boption\\b","\\buser\\b","\\bdrop\\b","\\bor\\b","\\bvalues\\b","\\bdump\\b","\\border\\b","\\bvarying\\b","\\belse\\b","\\bouter\\b","\\bview\\b","\\bend\\b","\\bover\\b", \
        "\\bwaitfor\\b","\\berrlvl\\b","\\bpercent\\b","\\bwhen\\b","\\bescape\\b","\\bpivot\\b","\\bwhere\\b","\\bexcept\\b","\\bplan\\b","\\bwhile\\b","\\bexec\\b","\\bprecision\\b","\\bwith\\b", \
        "\\bexecute\\b","\\bprimary\\b","\\bwithingroup\\b","\\bexists\\b","\\bprint\\b","\\bwritetext\\b","\\bexit\\b","\\bproc\\b"]

        self.highlightingRules = [(QRegExp(pattern), keywordFormat)  for pattern in keywordPatterns]



        funcFormat = QTextCharFormat()
        funcFormat.setForeground(Qt.darkCyan)
        funcFormat.setFontWeight(QFont.Bold)
        functionPatterns = ["\\UPPER\\b","\\ABS\\b","\\AVG\\b","\\CEILING\\b","\\COUNT\\b","\\FLOOR\\b","\\MAX\\b","\\MIN\\b","\\RAND\\b","\\ROUND\\b","\\SIGN\\b","\\SUM\\b", \
        "\\upper\\b","\\abs\\b","\\avg\\b","\\ceiling\\b","\\count\\b","\\floor\\b","\\max\\b","\\min\\b","\\rand\\b","\\round\\b","\\sign\\b","\\sum\\b"]
        [self.highlightingRules.append((QRegExp(pattern), funcFormat))  for pattern in functionPatterns]


        classFormat = QTextCharFormat()
        classFormat.setFontWeight(QFont.Bold)
        classFormat.setForeground(Qt.darkMagenta)
        self.highlightingRules.append((QRegExp("\\bQ[A-Za-z]+\\b"),classFormat))

        singleLineCommentFormat = QTextCharFormat()
        singleLineCommentFormat.setForeground(Qt.red)
        self.highlightingRules.append((QRegExp("//[^\n]*"), singleLineCommentFormat))

        self.multiLineCommentFormat = QTextCharFormat()
        self.multiLineCommentFormat.setForeground(Qt.red)

        quotationFormat = QTextCharFormat()
        quotationFormat.setForeground(Qt.darkGreen)
        self.highlightingRules.append((QRegExp("\".*\""), quotationFormat))

        functionFormat = QTextCharFormat()
        functionFormat.setFontItalic(True)
        functionFormat.setForeground(Qt.darkCyan )
        self.highlightingRules.append((QRegExp("\\b[A-Za-z0-9_]+(?=\\()"),functionFormat))              #((QRegExp("\\b[A-Za-z0-9_]+(?=\\()"),functionFormat))

        self.commentStartExpression = QRegExp("/\\*")
        self.commentEndExpression = QRegExp("\\*/")

    def highlightBlock(self, text):
        for pattern, format in self.highlightingRules:
            
            expression = QRegExp(pattern)
            index = expression.indexIn(text)
            while index >= 0:
                length = expression.matchedLength()
                self.setFormat(index, length, format)
                index = expression.indexIn(text, index + length)

        self.setCurrentBlockState(0)

        startIndex = 0
        if self.previousBlockState() != 1:
            startIndex = self.commentStartExpression.indexIn(text)

        while startIndex >= 0:
            endIndex = self.commentEndExpression.indexIn(text, startIndex)

            if endIndex == -1:
                self.setCurrentBlockState(1)
                commentLength = len(text) - startIndex
            else:
                commentLength = endIndex - startIndex + self.commentEndExpression.matchedLength()

            self.setFormat(startIndex, commentLength,
                    self.multiLineCommentFormat)
            startIndex = self.commentStartExpression.indexIn(text,
                    startIndex + commentLength);



# Handling the Excel file generation class
class PandasModel(QtCore.QAbstractTableModel): 
    def __init__(self, df = pd.DataFrame(), parent=None): 
        QtCore.QAbstractTableModel.__init__(self, parent=parent)
        self._df = df

    def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
        if role != QtCore.Qt.DisplayRole:
            return QtCore.QVariant()

        if orientation == QtCore.Qt.Horizontal:
            try:
                return self._df.columns.tolist()[section]
            except (IndexError, ):
                return QtCore.QVariant()
        elif orientation == QtCore.Qt.Vertical:
            try:
                # return self.df.index.tolist()
                return self._df.index.tolist()[section]
            except (IndexError, ):
                return QtCore.QVariant()

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if role != QtCore.Qt.DisplayRole:
            return QtCore.QVariant()

        if not index.isValid():
            return QtCore.QVariant()

        return QtCore.QVariant(str(self._df.ix[index.row(), index.column()]))

    def setData(self, index, value, role):
        row = self._df.index[index.row()]
        col = self._df.columns[index.column()]
        if hasattr(value, 'toPyObject'):
            # PyQt4 gets a QVariant
            value = value.toPyObject()
        else:
            # PySide gets an unicode
            dtype = self._df[col].dtype
            if dtype != object:
                value = None if value == '' else dtype.type(value)
        self._df.set_value(row, col, value)
        return True

    def rowCount(self, parent=QtCore.QModelIndex()): 
        
        return len(self._df.index)

    def columnCount(self, parent=QtCore.QModelIndex()): 
        return len(self._df.columns)
        

    def sort(self, column, order):
        colname = self._df.columns.tolist()[column]
        self.layoutAboutToBeChanged.emit()
        self._df.sort_values(colname, ascending= order == QtCore.Qt.AscendingOrder, inplace=True)
        self._df.reset_index(inplace=True, drop=True)
        self.layoutChanged.emit()


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    #HiLight = Highlighter()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
    

