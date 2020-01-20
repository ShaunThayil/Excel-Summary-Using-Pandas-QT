import pandas as pd
xl_file = pd.read_excel("Modified_Khardung.xlsx")

#xl_file = pd.read_excel("KH-PH-AWS-2017-19-Johnn.xlsx",sheet_name="Khardung-AWS-SEB")

# for i in range(909,980 + 1):
#     date = str(xl_file.loc[i,'DATE'])
#     new_str_date = date[1:3]+'-'+date[0]+'-20'+date[-2:]
#     xl_file.loc[i,'DATE'] = pd.to_datetime(new_str_date)


# xl_file = pd.read_excel("Modified_Khardung.xlsx")
# for i in range(1893,2468 + 1):
#     date = xl_file.loc[i,'DATE']
#     xl_file.loc[i,'DATE']=date.replace(month=date.day,day=date.month)

for i in range(909,3141):
    print(i)
    time = xl_file.loc[i,'TIME']
    str_time = str(time)
    if(time==0):
        new_time = '00:00:00'
    elif(time==30):
        new_time = '00:30:00'
    elif(time<1000):
        new_time = '0'+str_time[0]+':'+str_time[-2:]+':00'
    else:
        new_time = str_time[:-2]+':'+str_time[-2:]+':00'
    xl_file.loc[i,'TIME'] = new_time
    if(i==3139):
      xl_file.to_excel("Modified_Khardung_time.xlsx")

#xl_file.loc[3142,'TIME'] = "23:30:00"

#xl_file.to_excel("Modified_Khardung_time.xlsx")
