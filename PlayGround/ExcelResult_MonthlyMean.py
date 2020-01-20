# This Program is used to Extract The Mean,Maximum and Minimum Of Each Column In A Period Of 30 Days
# All The Computation Takes Place Inside The Class ExcelResult
# Class Variables:- file_name ,save_name , sheet_name ,target_workbook ,index_list , column_names
# Class Methods:-   get_result , daily_means_max_min , create_sheets , unique_index  , col_names


import datetime
import pandas as pd
import time
import sys
import calendar


class ExcelResult_MonthlyMean:
    def __init__(self, excel_file, savepath):
        #print("Hello Moto1")
        self.excel_file = excel_file  # Initialize The Class Excel Document
        #print("After Init1")
        self.sheet_names= self.excel_file.sheet_names # Get A List A Of Columns  In The Excel File Without Unwanted Columns (Function Return A List)
       #print("Hello Moto2")
        self.writer = pd.ExcelWriter(savepath, engine='xlsxwriter')  # Writer To The SavePath
        #print("Hello Moto2")
        self.get_result()  # Computer The Values And Write To The New Excel File
        #print("Hello Moto3")
        self.writer.save()  # Save The Excel File (Result Will Appear Only When You Save The File)
       # print("Hello Moto3")
    def get_result(self):
        for i in self.sheet_names:
            excel_sheet=self.excel_file.parse(i)
            self.monthly_means_max_min(excel_sheet,i)

    def monthly_means_max_min(self,xl_file,sheet_name):  # To Add Values To The New Sheet
        df = pd.DataFrame(columns=['Year_Month', 'MEAN', 'MAX', 'MIN'])  # Create A DataFrame To Store Data Temporarily
        i=0
        col_names=list(xl_file.columns.values)
        while(i<len(xl_file)-1):
            
            year=xl_file.loc[i, 'DATE'].year
            month=xl_file.loc[i,'DATE'].month
            month_days=calendar.monthrange(year,month)[1]
            mean = xl_file.loc[i:i+month_days-1,'MEAN'].mean()
            maximum = xl_file.loc[i:i+month_days-1, 'MAX'].max()
            minimum = xl_file.loc[i:i+month_days-1, 'MIN'].min()
            print(xl_file.loc[i,'DATE'],"-",xl_file.loc[i+month_days-1,'DATE'])
            year_month=str(year)+" "+ calendar.month_name[month]
            df.loc[len(df)] = [year_month, mean, maximum, minimum]  # Append values To The Temp_List
            i+=month_days
            


        df.to_excel(self.writer, sheet_name=sheet_name,index=False)
        print("Sheet Created",sheet_name)







def main():
    t1 = time.time()
    filename="Modified_Khardung_daily.xlsx"
    xl_file=pd.ExcelFile(filename)
    savename = "Modified_Khardung_monthly.xlsx"


    ExcelResult_MonthlyMean(xl_file, savename)
    t2 = time.time()
    print("TIME=", (t2 - t1))

main()