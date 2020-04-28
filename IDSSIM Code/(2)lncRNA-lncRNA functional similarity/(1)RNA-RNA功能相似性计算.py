import xlrd
import xlwt

#打开第1、2,3个矩阵
data1=xlrd.open_workbook('MNDR-lncRNA-disease association-meshname.xls')
data2=xlrd.open_workbook('MNDR-disease semantic similarity.xls')
data3=xlrd.open_workbook('MNDR-LNCRNA.xls')

#打开每个excel表中的第一个sheet工作表
table1=data1.sheets()[0]
row1=table1.nrows
col1=table1.ncols

table2=data2.sheets()[0]
row2=table2.nrows
col2=table2.ncols

table3=data3.sheets()[0]
row3=table3.nrows
col3=table3.ncols

matrix1=[[0 for i in range(col1)]for j in range(row1)]
matrix2=[[0 for i in range(col2)]for j in range(row2)]
matrix3=[[0 for i in range(row3)]for j in range(row3)]

for i in range(1,row3):
    matrix3[i][0] = table3.cell(i, 0).value
    matrix3[0][i]=matrix3[i][0]

#--------------------------------------------------------------
#定义疾病相似性函数
def similarDia(dia1,dia2):
    global values
    if(dia1==dia2):
        values=1
    else:
        for i in range(row2):
            if(dia1==table2.cell(i,0).value):
                for j in range(col2):
                    if(dia2==table2.cell(0,j).value):
                        values=table2.cell(i,j).value
    return values

#----------------------------------------------------------
#定义疾病组函数
def DiaRna(RNA):
    asso=[]
    for i in range(1,row1):
        if(RNA==table1.cell(i,0).value):
            asso.append(table1.cell(i,2).value)
    return asso

#------------------------------------------------------------
#定义RNA-RNA相似函数
def similarRNA(RNA1,RNA2):
    DvRNA1=DiaRna(RNA1)
    DvRNA2=DiaRna(RNA2)
    len1=len(DvRNA1)
    len2=len(DvRNA2)

    s=[[0 for i in range(len2)] for j in range(len1)]
    a=0
    b=0

    for i in DvRNA1:
        for j in DvRNA2:
            s[a][b]=similarDia(i,j)
            b=b+1
        a=a+1
        b=0
    sum1=0
    sum2=0
    sum = 0

    for m in range(len1):                             #将每行的最大值相加
       sum1+=(max(s[m][:]))

    for n in range(len2):                            #将每列的最大值相加
        sum = [x[n] for x in s]
        sum2 += (max(sum))
    return ((sum1+sum2)/(len1+len2))                 #返回RNA-RNA的相似性

#-----------------------------------------------------------------------------------
f = xlwt.Workbook()
sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet

for i in range(row3):
    for j in range(row3):
        if(i==0):
            sheet1.write(0,j,matrix3[0][j])
            print(0,j,matrix3[0][j])
        if(i!=0 and j==0):
            sheet1.write(i,0,matrix3[i][0])
            print(i, 0, matrix3[i][0])
        if(i!=0 and j!=0):
            sheet1.write(i,j,similarRNA(matrix3[i][0],matrix3[0][j]))
            print(i, j, similarRNA(matrix3[i][0], matrix3[0][j]))


f.save('MNDR-lncRNA fuctional similarity.xls')