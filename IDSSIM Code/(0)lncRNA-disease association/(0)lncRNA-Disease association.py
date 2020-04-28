import xlrd
import xlwt
import pandas as pd

data1=xlrd.open_workbook('MNDRv2.0-lncRNA-Disease association.xlsx')
table1=data1.sheets()[0]
row1=table1.nrows
col1=table1.ncols

data2=xlrd.open_workbook('MNDR-LNCRNA.xls')
table2=data2.sheets()[0]
row2=table2.nrows
col2=table2.ncols


#------------------------------------------------------------------------
#创建新的lncRNA-disease关联文件
book=xlwt.Workbook()
sheet=book.add_sheet('test',cell_overwrite_ok=True)

#筛选LncRNA-disease关联-包含重复关联
k=1
sheet.write(0, 0, 'lncRNA名称')
sheet.write(0, 1, 'lncRNA关联疾病')
for i in range(1,row2):
    for j in range(1,row1):
        if(table2.cell(i,0).value==table1.cell(j,2).value):
            sheet.write(k,0,table2.cell(i,0).value)
            sheet.write(k, 1, table1.cell(j, 5).value)
            k=k+1

book.save('MNDR-lncRNA-disease association.xls')

#------------------------------------------------------------------------
#去除重复的关联行信息
data = pd.DataFrame(pd.read_excel('MNDR-lncRNA-disease association.xls'))
# 查看去除重复行的数据
no_re_row = data.drop_duplicates()
# 将去除重复行的数据输出到excel表中
no_re_row.to_excel("MNDR-lncRNA-disease association.xls")

#------------------------------------------------------------------------






