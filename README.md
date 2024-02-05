import pandas as pd
import xlsxwriter
import numpy as np
import os
import xlrd
 
path = 'analyses_LG'
folder = [file for file in os.listdir(path)]
language_code = pd.read_excel('C:\\Users\\ilesanmi.amoran\\python\\internal_files\\internal_LG.xlsx', sheet_name='Data')
 
all_file = pd.DataFrame()  # Initialize an empty DataFrame to store concatenated data
 
for file in folder:
    df = pd.read_excel('./analyses_LG/' + file, sheet_name=0, header=1)
    df = df.drop(['File Name', 'Segments'], axis=1)
    df['Fuzzies'] = df['75% - 99%'] + df['Repetitions75% - 99%']
    df['New_Client'] = df['Repetitions75% - 99%'] + df['No Match']
 
    # Compare 'language' column in path file with language_code and add 'Support Code'
    df = pd.merge(df, language_code, how='left', on='Language')
 
    all_file = pd.concat([all_file, df], axis=0)
    
    Supplier_pivot = pd.pivot_table(all_file, index=['Language', 'Batch Name'],
                                values=('Context', '100% Match', 'Fuzzies', 'No Match', 'Repetitions', 'Total'),
                                aggfunc='sum').reset_index()
    Supplier_pivot = Supplier_pivot[['Language', 'Batch Name', 'Context', '100% Match', 'Fuzzies', 'No Match', 'Repetitions', 'Total']]
    Client_pivot = pd.pivot_table(all_file, index=['Language', 'Batch Name'],
                              values=('Context', '100% Match', '75% - 99%', 'New_Client', 'Repetitions', 'Total'),
                              aggfunc='sum').reset_index()
    Client_pivot = Client_pivot[['Language', 'Batch Name', 'Context', '100% Match', '75% - 99%', 'New_Client', 'Repetitions', 'Total']]
 
# Write to Excel
with pd.ExcelWriter('LG_file.xlsx', engine='xlsxwriter') as writer:
    Client_pivot.to_excel(writer, sheet_name='CLIENT', index=False)
    Supplier_pivot.to_excel(writer, sheet_name='SUPPLIER', index=False)
