import pandas as pd 
import time
import numpy as np
import datetime
import time


t1=time.time()
excel_file=pd.ExcelFile("AWS4 DATA 2012-14.xlsx")
sheetnames=list(excel_file.sheet_names)
excel_sheet=excel_file.parse(sheetnames[1])
new_df=excel_sheet['SunDur'].groupby(lambda x:x/24)
print(list(new_df))
t2=time.time()
print("TIME=",(t2-t1))
#for k,g in excel_sheet.groupby(np.arange(len(excel_sheet))//24):
 #print(k,g['SunDur'].mean())