import pandas as pd
import numpy as np
import jinja2
from openpyxl.styles import colors


sn1 = pd.ExcelFile('file1.xlsx').sheet_names
sn2 = pd.ExcelFile('file2.xlsx').sheet_names

df1 = pd.read_excel('file1.xlsx', sheet_name=None)
df2 = pd.read_excel('file2.xlsx', sheet_name=None)

writer = pd.ExcelWriter('./diff.xlsx', engine='xlsxwriter')





def writeDiffs(f1, f2, sheetName):
    print("-------------------------------------------")
    _dftmp = pd.read_excel('file1.xlsx', sheet_name=sheetName)

    print("sheet1=" + sheetName + " and sheet2=" + sheetName)

    comparison_values = f1[sheetName].values == f2[sheetName].values

    print(comparison_values)

    rows, cols = np.where(comparison_values == False)


    for item in zip(rows, cols):
        print()
        _dftmp.iloc[item[0], item[1]] = '{} --> {}'.format(f1[sheetName].iloc[item[0], item[1]],
                                                           f2[sheetName].iloc[item[0], item[1]])


        _dftmp.to_excel(writer, sheet_name=sheetName, index=False, header=False)


for sheet in sn1:
    writeDiffs(df1, df2, sheet)
    #writeDiffs(df1, df2, 'Sheet2')

writer.save()
print("Successfull")
writer.close()
