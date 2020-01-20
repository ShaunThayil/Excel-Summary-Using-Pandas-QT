#!usr/bin/python
from PyQt5 import QtGui,QtCore,QtWidgets,uic
import sys
    
import os
import pandas as pd
import ExcelResult_DailyMean
import ExcelResult_HourlyMean
import ExcelResult_MonthlyMean
class Window(QtWidgets.QMainWindow):                              #Window Class Inhertits From QMainWindow
    def __init__(self):                                       #Python Constructor
        super().__init__()                         #SuperClass Constructor Call
        self.ui=uic.loadUi("pa-gui.ui",self)


        self.ui.quit.clicked.connect(self.close_application) #Connects to The Function to be executed When The Option is Clicked

        #MENUBAR
        self.ui.actionQuit.triggered.connect(self.close_application)

        self.home()


    def home(self):

        


        #Button To Choose File

        self.ui.choose_file.clicked.connect(self.get_file_name)



        
         #DropDown To Choose Sheet
        self.ui.sheet_drop_down.activated[str].connect(self.choose_sheet)


        
        
        #Button To Choose Save Directory

        self.ui.save_folder.clicked.connect(self.get_dir_name)

        
        
        self.ui.hourly_mean.setChecked(True)



        
        



        #Button To Get Save Path

        self.ui.start_calc.clicked.connect(self.start)


        self.show()                                             #Show Everything
    
    #Func To Close
    def close_application(self):
        
        choice=QtWidgets.QMessageBox.question(self,"Quit","Do You Really Wanna Exit?",QtWidgets.QMessageBox.Yes|QtWidgets.QMessageBox.No)   #MessageBox
        if(choice ==QtWidgets.QMessageBox.Yes):
         sys.exit()
        else:
          pass
   
    def start(self):

        #save_filename=str(self.ui.save_filename.text())+".xlsx"
        #self.savename=os.path.join(self.dirPath,save_filename)
        #print "FilePath=",self.file_name
        #print "SheetName=",self.sheetname
        #print "SavePath=",self.savename,type(self.savename)
        if(self.ui.hourly_mean.isChecked() == True):
             print("Choose Hourly")

             ExcelResult_HourlyMean.ExcelResult_HourlyMean(self.excel_file.parse(str(self.sheetname)), self.savename)
        elif(self.ui.daily_mean.isChecked()== True):
            print("Choose Daily")
            ExcelResult_DailyMean.ExcelResult_DailyMean(self.excel_file.parse(str(self.sheetname)), self.savename)
        else:
            print("Choose Monthly")
            ExcelResult_MonthlyMean.ExcelResult_MonthlyMean(self.excel_file,self.savename)
        self.ui.completed.setText("Completed!!!")

    

    def get_file_name(self):
        self.ui.completed.setText("")
        file_name , _=QtWidgets.QFileDialog.getOpenFileName(self,"Open File")
        self.ui.filePath.setText(file_name)

        self.excel_file=pd.ExcelFile(file_name)
        sheetnames=self.excel_file.sheet_names
        print("GOT IT")
        self.ui.sheet_drop_down.clear()                   #To  Clear The List Each Time A New File Is Chosen
        self.sheetname=sheetnames[0]
        for i in sheetnames:
            print(i)
            self.ui.sheet_drop_down.addItem(i)          #Add SheetNames To The Combobox

    def choose_sheet(self,text):
      self.sheetname=text
    

    def get_dir_name(self):
       self.savename , _ = QtWidgets.QFileDialog.getSaveFileName(self, "Select Directory")
       self.ui.save_file_txt.setText(self.savename)
       
    def get_save_path(self):
        save_filename=str(self.ui.save_filename.text())
        #self.savename=os.path.join(self.dirPath,save_filename)    #To Join The String Using The File Separator Used In The OS
       


def run():
    app=QtWidgets.QApplication(sys.argv)                        #sys.argv is a list which contains the Command Line Arguments
    GUI=Window()                                            #Make a Window Object
    sys.exit(app.exec_())                                    #For Clean Exit Won't Run Without It

run()