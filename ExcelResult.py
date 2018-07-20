#This Program is used to Extract The Mean,Maximum and Minimum Of Each Column In A Period Of 30 Days
#All The Computation Takes Place Inside The Class ExcelResult 
#Class Variables:- file_name ,save_name , sheet_name ,target_workbook ,index_list , column_names
#Class Methods:-   get_result , daily_means_max_min , create_sheets , unique_index  , col_names


import datetime 
import pandas as pd
import time
import sys

class ExcelResult:
    def __init__(self,excel_file,savepath):
        self.excel_file=excel_file        #Initialize The Class Excel Variable
        self.column_names=self.col_names()     #Get A List A Of Columns  In The Excel File Without Unwanter Columns
        self.uniq_date=self.unique_date()    #Get A List Of Dates In Julian Date
        self.writer = pd.ExcelWriter(savepath, engine='xlsxwriter')  #Writer To The SavePath
        self.get_result()    #Computer The Values And Write To The New Excel File
        self.writer.save()   #Save The Excel File (Result Will Appear Only When You Save The File)


    def get_result(self):
        for i in self.column_names:
            self.daily_means_max_min(i)
    
    def daily_means_max_min(self,col_name): #To Add Values To The New Sheet
        df=pd.DataFrame(columns=['DATE','MEAN','MAX','MIN'])         #Create A DataFrame To Store Data Temporarily
        j=0
        for i in range(0,len(self.excel_file),24):
             mean=self.excel_file.loc[i:i+23,col_name].mean()
             maximum=self.excel_file.loc[i:i+23,col_name].max()
             minimum=self.excel_file.loc[i:i+23,col_name].min()
             df.loc[len(df)]=[self.uniq_date[j],mean,maximum,minimum]   #Append values To The Temp_List
             j+=1
             ##df['DATE']=self.uniq_date
             ##df['MEAN']=self.new_gp[col_name].mean()
             ## df['MAX']=self.new_gp[col_name].max()
             ## df['MIN']=self.new_gp[col_name].min()
      
      
        df.to_excel(self.writer, sheet_name=col_name)
        print("Sheet Created",col_name)
    
    
    def unique_date(self):    #Returns a list of unique dates In The Form Of Julian Dates
        list1=list(self.excel_file['DATE'].unique())
        list2=[]
        for i in list1:
           k=pd.Timestamp(i).to_julian_date()   #In Built Function To Convert Timestamps To Julian Date
           list2.append(k)
        return list2

    def col_names(self):   #Returns The Column Header Names Without Unwanted Columns
        col_name=list(self.excel_file.columns.values)
        col_name.remove("DATE")    #Removes Date And Time As They Are Not Need In Calculations
        col_name.remove("TIME")
        return col_name
    

    #def check(self,cell,ls):
     #if(type(cell)!=datetime.time):
      # ls.append(cell.time())
     #else:
     #   ls.append(cell)


    #Convert Datetime Cells in Time Field To datetime.time Fields
    #def uniform_datatype(self):
     #  ls=[]
      # self.excel_sheet.apply(lambda row:self.check(row['TIME']),axis=1)
     #  self.excel_sheet['TIME']=ls



def main():
    t1=time.time()
    filename="AWS4 DATA 2012-14.xlsx"  #The File You Need To Process
    sheetname="Ist process"
    savename="AWS4-Data-2012-2014_DailyMeans.xlsx"
    xl_file=pd.read_excel(filename,sheet_name=sheetname)

    ExcelResult(xl_file,savename)
    t2=time.time()
    print("TIME=",(t2-t1))

