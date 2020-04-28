import xlrd
import xlwt
import copy
import math
import numpy as np

datamesh=xlrd.open_workbook('meshtree.xlsx')
tablemesh=datamesh.sheets()[0]
rowmesh=tablemesh.nrows


dataDAG=xlrd.open_workbook('Disease probability in mesh.xls')
tableDAG=dataDAG.sheets()[0]
rowDAG=tableDAG.nrows

datapath=xlrd.open_workbook('MNDR-disease meshpath.xls')
tablepath=datapath.sheets()[0]
rowpath=tablepath.nrows

#-------------------------------------------------------------------
#构建一个疾病的所有路径矩阵
def matrixpath(disease):
    row1=0
    col1=0

    for i in range(rowpath):
        if(tablepath.cell(i,1).value==disease):
            row1+=1
            lenth=len(tablepath.cell(i,0).value)
            sum=math.ceil(lenth/4)
            if(col1<sum):
                col1=sum
    x=0
    y=0
    matrixpth = [[0 for i in range(col1)] for j in range(row1)]
    for i in range(rowpath):
        if (tablepath.cell(i, 1).value == disease):
            num=math.ceil((len(tablepath.cell(i,0).value)/4))
            for j in range(i,i+num):
                matrixpth[x][y]=tablepath.cell(j,1).value
                y+=1
            x+=1
            y=0


    return(matrixpth)

#----------------------------------------------------------------------
#构造二维数组，第一列为不重复的所有节点，第二列为所在层数
def sumcode(path):
    row=len(path)
    col=len(path[0])
    matrix=copy.deepcopy(path)
    list = []
    for i in range(row - 1):
        for j in range(col):
            for k in range(i + 1, row):
                for l in range(col):
                    if (matrix[i][j] == matrix[k][l]):
                        matrix[k][l] = 0

    for i in range(row):
        for j in range(col):
            if (matrix[i][j] != 0):
                list.append(matrix[i][j])

    lenlist = len(list)
    list1 = [[0 for i in range(2)] for j in range(lenlist)]
    for i in range(lenlist):
        for j in range(2):
            if (j == 0):
                list1[i][j] = list[i]
            else:
                list1[i][j] = 0

    return list1

#------------------------------------------------------------------
#在疾病路径矩阵中，去除重复节点
def dletsimivode(disease):
    diamatrix=matrixpath(disease)
    row = len(diamatrix)
    col = len(diamatrix[0])
    diamatrix2 = [[0 for i in range(col)] for j in range(row)]
    for i in range(row):
        for j in range(col):
            diamatrix2[i][j] = j
    diamatrix3 = sumcode(diamatrix)
    row3 = len(diamatrix3)

    sum = 0
    for i in range(row3):
        for k in range(row):
            for l in range(col):
                if (diamatrix[k][l] == diamatrix3[i][0]):
                    sum = sum + 1
                    if (sum == 1):
                        min = diamatrix2[k][l]
                    else:
                        if (min > diamatrix2[k][l]):
                            min = diamatrix2[k][l]
        diamatrix3[i][1] = min

        for m in range(row):
            for n in range(col):
                if (diamatrix3[i][0] == diamatrix[m][n]):
                    if (diamatrix2[m][n] > min):
                        temp = diamatrix2[m][n] - min
                        for p in range(n, col):
                            diamatrix2[m][p] = diamatrix2[m][p] - temp
            temp = 0
        sum = 0
        min = 0
    return diamatrix3

#----------------------------------------------------------------------------------
def childpvalue(disease):
    for i in range(rowDAG):
        if(disease==tableDAG.cell(i,0).value):
            return tableDAG.cell(i,1).value

#--------------------------------------------------------------------
#计算语义贡献因子 语义贡献因子用factor表示
def calfactor(disease):
    p=childpvalue(disease)
    c=1
    max=0.13022752628554
    factor=0.5+(max-p)*c

    return factor

#-----------------------------------------------------------------
#计算疾病术语的语义贡献值
def DVdisease(disease):
    pathmatrix=matrixpath(disease)       #所有路径矩阵
    dlesimatrix=dletsimivode(disease)
    row1=len(pathmatrix)
    col1=len(pathmatrix[0])
    row2=len(dlesimatrix)
    col2=len(dlesimatrix[0])

    #添加一列语义贡献因子
    for i in range(row2):
        dlesimatrix[i].append(calfactor(dlesimatrix[i][0]))
        dlesimatrix[i].append(0)

    dlesimatrix=sorted(dlesimatrix, key=lambda dlesimatrix: dlesimatrix[1])  # sort by 层数

    # 添加一列每个术语的语义贡献值
    sum = 0
    childlist = []  # 记录孩子节点在dlesimatrix的语义贡献值
    child = []      #记录原路径中的孩子节点
    lastchild=[]    #孩子节点的最终矩阵
    num = []
    for i in range(row2):
        if (i == 0):
            dlesimatrix[i][3]=1
        else:
            for j in range(row2):
                if (dlesimatrix[j][1] == dlesimatrix[i][1] - 1):
                    sum += 1
                    num.append(dlesimatrix[j][0])
                    num.append(dlesimatrix[j][3])
                    num.append(j)
                    childlist.append(num)
                    num = []
            childlist = np.array(childlist)
            if (sum == 1):
                dlesimatrix[i][3]=dlesimatrix[i][2] * dlesimatrix[int(childlist[0][2])][3]
            else:
                for k in range(row1):
                    for l in range(col1):
                        if (dlesimatrix[i][0] == pathmatrix[k][l]):
                            if pathmatrix[k][l - 1] not in child:
                                child.append(pathmatrix[k][l - 1])
                for m in range(sum):
                    if(childlist[m][0] in child):
                        lastchild.append(childlist[m][1])
                maxsemantic=max(lastchild)
                dlesimatrix[i][3]=dlesimatrix[i][2]*float(maxsemantic)

            #print(child)
            #print(childlist1)
            sum=0
            child=[]
            childlist=[]
            lastchild=[]
    return (dlesimatrix)

#-----------------------------------------------
#计算每个疾病的语义贡献值
def DV(disease):
    diamatrix=DVdisease(disease)
    row=len(diamatrix)
    dv=0
    for i in range(row):
        dv=dv+diamatrix[i][3]
    return dv

#-----------------------------------------------
#计算疾病-疾病之间的相似性

def similardia(disease1,disease2):
    if(disease1==disease2):
        result=1
    else:
        dia1 = DVdisease(disease1)
        dia2 = DVdisease(disease2)
        row1 = len(dia1)
        row2 = len(dia2)
        sum = 0
        for i in range(row1):
            for j in range(row2):
                if(dia1[i][0]==dia2[j][0]):
                    sum=sum+dia1[i][3]+dia2[j][3]
                    break
        dv1 = DV(disease1)
        dv2 = DV(disease2)
        result = sum / (dv1 + dv2)

    return result

#---------------------------------------------------------------------
data2=xlrd.open_workbook('MNDR-associated total diseases.xls')
table2=data2.sheets()[0]
rowdia=table2.nrows
coldia=table2.ncols

f = xlwt.Workbook()
sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)

for i in range(rowdia):
 for j in range(rowdia):
     if(i==0):
         sheet1.write(0,j,table2.cell(j,0).value)
         print(0,j,table2.cell(j,0).value)
     if(i!=0 and j==0):
         sheet1.write(i,0,table2.cell(i,0).value)
         print(i,0,table2.cell(i, 0).value)
     if(i!=0 and j!=0):
         sheet1.write(i, j, similardia(table2.cell(i, 0).value, table2.cell(j, 0).value))
         print(i, j, ' ', similardia(table2.cell(i, 0).value, table2.cell(j, 0).value))


f.save('MNDR-disease semantic similarity.xls')