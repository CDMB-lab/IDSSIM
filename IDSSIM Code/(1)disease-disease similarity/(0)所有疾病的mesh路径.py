import xlrd
import xlwt
import copy
import numpy as np

datamesh=xlrd.open_workbook('meshtree.xlsx')
tablemesh=datamesh.sheets()[0]
rowmesh=tablemesh.nrows

data2=xlrd.open_workbook('MNDR-associated total diseases.xls')
table2=data2.sheets()[0]
rowdia=table2.nrows
coldia=table2.ncols

f = xlwt.Workbook()
sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)


sumcol=0

def matrixpath(disease):
    global sumcol
    for i in range(rowmesh):
        if(tablemesh.cell(i,1).value==disease):
            sheet1.write(sumcol,0,tablemesh.cell(i,0).value)
            sheet1.write(sumcol, 1, tablemesh.cell(i, 1).value)
            sumcol=sumcol+1
            lenth=len(tablemesh.cell(i,0).value)
            diavalue=tablemesh.cell(i, 0).value
            while (lenth > 3):
                for j in range(i):
                    if (tablemesh.cell(j, 0).value == diavalue[:-4]):
                        sheet1.write(sumcol,0,tablemesh.cell(j, 0).value)
                        sheet1.write(sumcol, 1, tablemesh.cell(j, 1).value)
                        sumcol+=1
                        diavalue = tablemesh.cell(j, 0).value
                        lenth = lenth - 4



for i in range(1,rowdia):
    matrixpath(table2.cell(i,0).value)


f.save('MNDR-disease meshpath.xls')