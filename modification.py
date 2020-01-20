# PreProcessing :  Creating Consistent Formats For Date Column

import pandas as pd
filename = "Modified_Phuche.xlsx"  # The File You Need To Process


xl_file = pd.read_excel(filename)

#xl_file['DATE'] = pd.to_datetime(xl_file['DATE'])
for i in range(11649,12224+1):
    xl_file.loc[i,'DATE'] = xl_file.loc[i,'DATE'].replace(day = xl_file.loc[i,'DATE'].month,month=xl_file.loc[i,'DATE'].day) # Change Date Format

print(xl_file.loc[11645:11649,'DATE'])
print(xl_file.loc[12220:12230,'DATE'])

xl_file.to_excel('Modified_Phuche.xlsx',sheet_name='Phuche')