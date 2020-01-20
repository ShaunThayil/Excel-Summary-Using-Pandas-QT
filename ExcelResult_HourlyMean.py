# This Program is used to Extract The Mean,Maximum and Minimum Of Each Column In A Period Of 30 Days
# All The Computation Takes Place Inside The Class ExcelResult
# Class Variables:- file_name ,save_name , sheet_name ,target_workbook ,index_list , column_names
# Class Methods:-   get_result , daily_means_max_min , create_sheets , unique_index  , reduced_col_names


import datetime
import pandas as pd
import time
import sys
import numpy as np



class ExcelResult_HourlyMean:
    def __init__(self, excel_file, savepath):
        print("Executing...")
        self.excel_file = excel_file                                            # Initialize The Class Excel Variable
        #self.excel_file['DATE'] = pd.to_datetime(self.excel_file['DATE'])  # if the data contains string date convert to datetime
        self.writer = pd.ExcelWriter(savepath, engine='xlsxwriter')             # Writer To The SavePath
        self.all_columns=list(excel_file.columns.values)
        self.calc_col_names=self.reduced_col_names()                                      #Get A List A Of Columns  In The Excel File Without Unwanted Columns
        self.preprocessing()
        self.excel_file=self.excel_file.interpolate(method='linear')


        self.save_sheetname = savepath.split('/')[-1].split('.')[0]                    # Saving Sheet-Name With The Name Of File Name(Applicable only in windows)
        print(self.calc_col_names)
        self.hourly_mean()                                                      # Computer The Values And Write To The New Excel File




    def hourly_mean(self):  # To Calculate Mean And Add  To The New Sheet
        print("Hourly_Mean")
        df=pd.DataFrame(columns=self.all_columns)
        xl_file=self.excel_file
        for i in range(0, len(xl_file)-1, 2):
            ls=[]
            #if(type(xl_file.loc[0,'DATE'])==pandas._libs.tslibs.timestamps.Timestamp):
            ls.append(xl_file.loc[i+1,'DATE'].date())
            ls.append(xl_file.loc[i+1,'TIME'])
            print("Row -> ",i,"\n")
            for j in self.calc_col_names:
              mean =xl_file.loc[i:i + 1, j].mean()
              ls.append(mean)

            df.loc[len(df)] = ls  # Append new tuple To The Temporary DataFrame


        df.to_excel(self.writer, sheet_name=self.save_sheetname)  #Write The Temp Dataframe To Excel File
        self.writer.save()                                        # Save The Excel File (Result Will Appear Only When You Save The File)
        print("Sheet Created", self.save_sheetname)
        print("COMPLETED!!!")
    
    
    # Perform Necessary Preprocessing
    def preprocessing(self): 
        temp_df=self.excel_file.replace('NAN',np.nan)
        temp_df=temp_df.replace('INF',np.nan)
        temp_df = temp_df.replace('-.-',np.nan)
        temp_df.interpolate(method='linear')
        self.excel_file=temp_df

     # Returns The Column  Names Without Unwanted Columns
    def reduced_col_names(self): 
        col_name = list(self.excel_file.columns.values)
        rem_col = ['DATE','TIME']
        for col in rem_col:
            col_name.remove(col)
        return col_name



def main():
    t1 = time.time()
    filename = "Modified_Khardung.xlsx"  # The File You Need To Process

    savename = "Modified_Khardung_hourly.xlsx"
    xl_file = pd.read_excel(filename)

    # print(xl_file.head())

    #xl_file['DATE']=pd.to_datetime(xl_file['DATE'])
    # print(type(xl_file.loc[10737,'DATE']))
    # print(type(xl_file.loc[10736,'DATE']))
    

    ExcelResult_HourlyMean(xl_file, savename)
     
    t2 = time.time()
    print("TIME=", (t2 - t1))

main()
