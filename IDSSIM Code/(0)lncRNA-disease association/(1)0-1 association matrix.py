import xlrd
import xlwt

data1=xlrd.open_workbook('MNDR-lncrna-disease association-meshname.xls')
table1=data1.sheets()[0]
row1=table1.nrows
col1=table1.ncols

data2=xlrd.open_workbook('MNDR-LNCRNA.xls')
table2=data2.sheets()[0]
row2=table2.nrows
col2=table2.ncols

book1=xlwt.Workbook()
sheet1=book1.add_sheet('test',cell_overwrite_ok=True)

book2=xlwt.Workbook()
sheet2=book2.add_sheet('test',cell_overwrite_ok=True)

#--------------------------------------------------------------
#输出关联的所有疾病，为disease-disease相似性做准备
disecol=[]
j=1
for i in range(1,row1):
    if(table1.cell(i,2).value not in disecol):
        disecol.append(table1.cell(i,2).value)
        sheet1.write(j,0,table1.cell(i,2).value)
        j=j+1
book1.save('MNDR-associated total diseases.xls')

#----------------------------------------------------------------
#j为列长
matrix=[[0 for i in range(j)]for k in range(row2)]

for i in range(row2):
    for k in range(j-1):
        if(i==0):
            matrix[i][k+1]=disecol[k]
        if(i!=0 and k==0):
            matrix[i][0]=table2.cell(i,0).value


for m in range(1,row1):
    for x in range(1,row2):
        if(table1.cell(m,0).value==matrix[x][0]):
            for y in range(1,j):
                if(table1.cell(m,2).value==matrix[0][y]):
                    matrix[x][y]=1


for x in range(row2):
    for y in range(j):
        sheet2.write(x,y,matrix[x][y])

book2.save('MNDR-lncRNA-disease association matrix.xls')


