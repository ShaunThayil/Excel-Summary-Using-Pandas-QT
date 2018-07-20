from PyQt4 import QtGui,QtCore
import sys
    
import os
import pandas as pd
import ExcelResult
class Window(QtGui.QMainWindow):                              #Window Class Inhertits From QMainWindow
    def __init__(self):                                       #Python Constructor
        super(Window,self).__init__()                         #SuperClass Constructor Call
        self.setGeometry(100,100,600,400)                     #Set The Size Of The Window
        self.setWindowTitle("ExSum")                          #Set Title Of The Window
        

        extractAction=QtGui.QAction("Quit",self)   #Create An Item Of An Option In Menu Bar
        extractAction.setShortcut("Ctrl+Q")                   #Set The Shortcut To Invoke The Action no Space InBetween
        extractAction.setStatusTip("Quit The Window")     #Show Text In Status Bar When Mouse Tip Is Hovered over The icon
        extractAction.triggered.connect(self.close_application) #Connects to The Function to be executed When The Option is Clicked

          
        #MENUBAR
        mainMenu=self.menuBar()             #Initialize MenuBar
        fileMenu=mainMenu.addMenu("&File")  #Add An Option To The MenuBar
        fileMenu.addAction(extractAction)   #Add Command inside the dropdown list

      
        
       #Show Status Bar
        self.statusBar()     
        
        


        self.home()
    def home(self):                                                 
        #Label For Choose File 
        self.chooseFile=QtGui.QLabel("Choose File",self)
        self.chooseFile.move(80,30)
        
        #TextField To Show The File Path
        self.filePath=QtGui.QLineEdit(self)
        self.filePath.move(20,70)
        self.filePath.resize(250,20)
        self.filePath.setText("Click Button To Choose File")

        #Button To Choose File
        chooseFile_btn=QtGui.QPushButton("Open File",self)
        chooseFile_btn.clicked.connect(self.get_file_name)
        chooseFile_btn.setStatusTip("Open File") 
        chooseFile_btn.resize(chooseFile_btn.minimumSizeHint())
        chooseFile_btn.move(80,100)

        #Label To Choose Sheet
        sheetLabel=QtGui.QLabel("Choose Sheet",self)
        sheetLabel.move(400,30)
        
         #DropDown To Choose Sheet
        self.sheetDropDown=QtGui.QComboBox(self)
        self.sheetDropDown.resize(200,20)
        self.sheetDropDown.move(350,70)
        self.sheetDropDown.activated[str].connect(self.choose_sheet)

        #Label To Choose Save  Directory
        dirLabel=QtGui.QLabel("Choose Save Folder",self)
        dirLabel.resize(150,30)
        dirLabel.move(60,150)
        
        
        #Button To Choose Save Directory
        chooseDir_btn=QtGui.QPushButton("Choose Folder",self)
        chooseDir_btn.clicked.connect(self.get_dir_name)
        chooseDir_btn.setStatusTip("Choose Folder") 
        chooseDir_btn.resize(chooseDir_btn.minimumSizeHint())
        chooseDir_btn.move(70,180)
        
        
        
        #Label To Enter Save Filename
        saveLabel=QtGui.QLabel("Enter Save FileName",self)
        saveLabel.resize(150,30)
        saveLabel.move(380,150)

        #TextField To Enter Save FileName
        self.saveFile=QtGui.QLineEdit(self)
        self.saveFile.resize(150,20)
        self.saveFile.move(370,180)
        
        


          #PROGRESS BAR
        self.progress = QtGui.QProgressBar(self) 
        self.progress.setGeometry(200, 220, 250,100)

        #Button To Get Save Path
        getPath_btn=QtGui.QPushButton("START",self)
        getPath_btn.resize(getPath_btn.minimumSizeHint())
        getPath_btn.move(260,290)
        getPath_btn.clicked.connect(self.start)

        #Label To Indicate Process Completed
        self.compLabel=QtGui.QLabel("",self)
        self.compLabel.move(270,330)
        



       
        

        #Button To Quit
        btn=QtGui.QPushButton("Quit",self)
        btn.clicked.connect(self.close_application)             #Method To Execute When Clicked
        btn.resize(btn.minimumSizeHint())                              #Set Size Of The Button(Can BE (x,y) or minimumSizeHint() or sizeHint())
        btn.move(500,350)                                           #Set Location of the Button
       
        self.show()                                             #Show Everything
    
    #Func To Close
    def close_application(self):
        
        choice=QtGui.QMessageBox.question(self,"Quit","Do You Really Wanna Exit?",QtGui.QMessageBox.Yes|QtGui.QMessageBox.No)   #MessageBox
        if(choice ==QtGui.QMessageBox.Yes):
         sys.exit()
        else:
          pass
   
    def start(self):
        self.completed = 0
        save_filename=str(self.saveFile.text())+".xlsx"
        self.savename=os.path.join(self.dirPath,save_filename)
        #print "FilePath=",self.file_name
        #print "SheetName=",self.sheetname
        #print "SavePath=",self.savename,type(self.savename)
        
        ExcelResult.ExcelResult(self.excel_file.parse(str(self.sheetname)),self.savename) 
        self.compLabel.setText("Completed!!!")
    

    def get_file_name(self):
        file_name=str(QtGui.QFileDialog.getOpenFileName(self,"Choose File"))
        self.filePath.setText(file_name)
        
        self.excel_file=pd.ExcelFile(file_name,on_demand=True)
        sheetnames=self.excel_file.sheet_names
        
        self.sheetDropDown.clear()                   #To  Clear The List Each Time A New File Is Chosen
        self.sheetname=sheetnames[0]
        for i in sheetnames:
            self.sheetDropDown.addItem(i)          #Add SheetNames To The Combobox
    
    def choose_sheet(self,text):
      self.sheetname=text
    

    def get_dir_name(self):
       self.dirPath = str(QtGui.QFileDialog.getExistingDirectory(self, "Select Directory"))
       
    def get_save_path(self):
        save_filename=str(self.saveFile.text())
        self.savename=os.path.join(self.dirPath,save_filename)    #To Join The String Using The File Separator Used In The OS
       
        print(self.savename)

def run():
    app=QtGui.QApplication(sys.argv)                        #sys.argv is a list which contains the Command Line Arguments
    GUI=Window()                                            #Make a Window Object
    sys.exit(app.exec_())                                    #For Clean Exit Won't Run Without It

run()