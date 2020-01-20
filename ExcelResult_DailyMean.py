#This Program is used to Extract The Mean,Maximum and Minimum Of Each Column In A Period Of 30 Days
#All The Computation Takes Place Inside The Class ExcelResult 
#Class Variables:- file_name ,save_name , sheet_name ,target_workbook ,index_list , column_names
#Class Methods:-   get_result , daily_means_max_min , create_sheets , unique_index  , col_names


import datetime 
import pandas as pd
import time
import sys

class ExcelResult_DailyMean:
    def __init__(self,excel_file,savepath):


        self.excel_file=excel_file
        self.column_names=self.col_names()     
        print(self.column_names)
        self.col_len=len(self.column_names)
        print(self.excel_file.head())
        self.uniq_date=self.unique_date()
        print(len(self.uniq_date))
        self.writer = pd.ExcelWriter(savepath, engine='xlsxwriter')  #Writer To The SavePath
        self.get_result()   
        self.writer.save()   #Save The Excel File (Result Will Appear Only When You Save The File)

     #Computer The Values for each column And Write To The New Excel File
    def get_result(self):
        for i in self.column_names:
            self.daily_means_max_min(i)
    
    #Perfom Functions
    def daily_means_max_min(self,col_name): #To Add Values To The New Sheet
        df=pd.DataFrame(columns=['DATE','JULIAN DATE','MEAN','MAX','MIN'])         #Create A DataFrame To Store Data Temporarily
        j=0

        for i in range(0,len(self.excel_file),24):
             if(j==len(self.uniq_date)):
                break
             select_rows = self.excel_file.loc[i:i + 23,col_name]
             print(i)
             mean=select_rows.mean()
             maximum=select_rows.max()
             minimum=select_rows.min()
             #df.loc[len(df)]=[1,2,3,4]
             #print("success")
             #df.loc[len(df)]=[self.uniq_date[j][0],self.uniq_date[j][1],mean,maximum,minimum]   #Append values To The Temp_List
             print('j=',j,'DATE:',self.uniq_date[j][0],' JULIAN DATE:',self.uniq_date[j][1])
             df = df.append({'DATE':self.uniq_date[j][0],'JULIAN DATE':self.uniq_date[j][1],'MEAN':mean,'MAX':maximum,'MIN':minimum},ignore_index=True)
             j+=1
             


        self.write_data_to_xl(df,col_name)
      
     #To Write Data To The Excel Sheet
    def write_data_to_xl(self,df,col_name):
        print("Sheet To Be Created - ",col_name)
        df.to_excel(self.writer, sheet_name=col_name)
        print("Sheet Created", col_name)
    

     #Returns a list of unique dates In The Form Of (Gregorian Date,Julian Date) pairs
    def unique_date(self):   
        list1=list(self.excel_file['DATE'].unique())
        list2=[]
        for i in list1:
           ts=pd.Timestamp(i)
           gregorian_date=datetime.datetime(ts.year,ts.month,ts.day).date()   #Convert np.datetime64 to datetime.date
           julian_date=int(pd.Timestamp(i).to_julian_date())   #In Built Function To Convert Timestamps To Julian Date
           list2.append([gregorian_date,julian_date])
        
        return list2

    #Returns The Column Header Names Without Unwanted Columns
    def col_names(self):   
        col_name=list(self.excel_file.columns.values)
        rm_col = ['Unnamed: 0','DATE','TIME'] # Add the unwanter column names to this list
        for col in rm_col:
           if(col in col_name):
            col_name.remove(col)
        
        return col_name
    



# Append main() to the end of the code for standalone execution
def main():
    t1=time.time()
    filename="Modified_Khardung_hourly.xlsx"  #The File You Need To Process
  
    savename="Modified_Khardung_daily.xlsx"
    xl_file=pd.read_excel(filename)

    ExcelResult_DailyMean(xl_file,savename)
    t2=time.time()
    print("TIME=",(t2-t1))

