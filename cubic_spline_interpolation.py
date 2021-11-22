import xlrd
import pandas as pd
import scipy.interpolate as scipl
import numpy as np

#inport Excel file
WorkBook = xlrd.open_workbook('hoge\source.xlsx')

#select work sheet
Sheet_A = WorkBook.sheet_by_name('RawDataA')

#get last No. of rows/columns
LastOfRow = Sheet_A.nrows
LastOfCol = Sheet_A.ncols

#define List
Time = [] 
ID = []
ExportList = []

#Get Time Column
for i in range(4,1214):
    Time.append(Sheet_A.cell_value(i,0))

ExTime = np.arange(0,22000,10) #make arithmetic progression

#get LTdata & execute cubic spline
for i in range(1,LastOfCol):
    ID.append(Sheet_A.cell_value(3,i))
    LT = []
    for j in range(4,1214):
        LT.append(Sheet_A.cell_value(j,i))

    spl = scipl.CubicSpline(Time, LT) #spline
    ExportList.append(spl(ExTime))
    
#print(LT)
#print(ExportList)
#print(ID)

df = pd.DataFrame(ExportList, columns=ExTime, index=ID) #make dataframe
df = df.transpose() #transpose the dataframe
#print(df)

df_s = df.sort_index(axis=1) #Sort by ID
#print(df_s)

df_s.to_excel('Results.xlsx', sheet_name='A') #Create Result file
