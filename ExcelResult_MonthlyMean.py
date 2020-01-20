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
        self.excel_file = excel_file  # Initialize The Class Excel Document
        self.sheet_names= self.excel_file.sheet_names # Get A List A Of Columns  In The Excel File Without Unwanted Columns (Function Return A List)
        self.writer = pd.ExcelWriter(savepath, engine='xlsxwriter')  # Writer To The user provided SavePath
        self.get_result()  
        self.writer.save()  # Save The Excel File (Result Will Appear Only When You Save The File)

    # Computer The Values And Write To The New Excel Sheet
    def get_result(self):
        for i in self.sheet_names:
            excel_sheet=self.excel_file.parse(i)  # Access the Sheet corresponding to column i
            self.monthly_means_max_min(excel_sheet,i)

   
     # Computer Mean,Max,Min for each month for a column and add to the dataframe
    def monthly_means_max_min(self,xl_file,sheet_name):
        df = pd.DataFrame(columns=['Year_Month', 'MEAN', 'MAX', 'MIN'])  # Create A DataFrame To Store Data Temporarily
        i=0
        col_names=list(xl_file.columns.values)
        while(i<len(xl_file)-1):
            
            year=xl_file.loc[i, 'DATE'].year
            month=xl_file.loc[i,'DATE'].month
            month_days=calendar.monthrange(year,month)[1]           # Find the number of days in a month in a given year

            # Calculates the mean,max,min from 1st of a month till the last date of that month
            mean = xl_file.loc[i:i+month_days-1,'MEAN'].mean()          
            maximum = xl_file.loc[i:i+month_days-1, 'MAX'].max()
            minimum = xl_file.loc[i:i+month_days-1, 'MIN'].min()
            print(xl_file.loc[i,'DATE'],"-",xl_file.loc[i+month_days-1,'DATE']) # Info
            year_month=str(year)+" "+ calendar.month_name[month]    # Form the Month Year Pair
            df.loc[len(df)] = [year_month, mean, maximum, minimum]  # Append values To The Dataframe
            i+=month_days
            


        df.to_excel(self.writer, sheet_name=sheet_name,index=False)  # Save the DataFrame
        print("Sheet Created",sheet_name)






# Append main() to execute this function in a standalone manner
# df.loc is very important function which provides index based access to the database
def main():
    t1 = time.time()
    filename="Modified_Khardung_daily.xlsx"   # The File to be accesses
    xl_file=pd.ExcelFile(filename)
    savename = "Modified_Khardung_monthly.xlsx"


    ExcelResult_MonthlyMean(xl_file, savename)
    t2 = time.time()
    print("TIME=", (t2 - t1))  # Determine the Running Time

