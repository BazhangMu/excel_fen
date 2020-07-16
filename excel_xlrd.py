import xlrd as xr
import xlwt as xw
data = xr.open_workbook('整改台账.xls')
table3 = data.sheet_by_name(u'1')#通过名称获取
#table3 = data.sheet_by_index(1)
a=[]#所有数据
ll=int(input('按第几列拆分：'))
ll=ll-1#按第几列拆分 0
#kk=input('请输入分列依据:')
for i in range(table3.nrows):
    a +=[table3.row_values(i)]
bb =list(table3.col_values(ll))#按第ll列拆分
b=[]#分表字典键值
col=len(table3.row_values(0))
print(col)
for i in range(1,len(bb)):
    if bb[i] not in b:
        b.append(bb[i])
c={}#各分表数据字典
for i in range(len(b)):
    k=0
    c[b[i]]=[a[0]]
    for j in range(len(a)):
        if b[i]==a[j][ll]:  #按第ll列拆分
            c[b[i]] +=[a[j]]
for i in list(c.keys()):
    f=xw.Workbook()
    sheet1=f.add_sheet('1')
    for j in range(len(c[i])):# j 行
        if len(c[i])>1:
            for k in range(col):#k 列
                sheet1.write(j,k,c[i][j][k])
    f.save(i.strip('\t')+'.xls')
print('已生成文件：',b)
'''
#print(a)
#print(bb)
#print(b)
#print(c)
#print("333333",c['兴隆'])
table1 = data.sheets()[0]          #通过索引顺序获取
table2 = data.sheet_by_index(0) #通过索引顺序获取
'''
''' import 的不同使用方法
from xlrd import *
from xlwt import *
data = open_workbook('整改台账.xls')
table3 = data.sheet_by_name(u'1')#通过名称获取
a=[]#所有数据
ll=int(input('按第几列拆分：'))
ll=ll-1#按第几列拆分 0
#kk=input('请输入分列依据:')
for i in range(table3.nrows):
    a +=[table3.row_values(i)]
bb =list(table3.col_values(ll))#按第ll列拆分
b=[]#分表字典键值
col=len(table3.row_values(0))
print(col)
for i in range(1,len(bb)):
    if bb[i] not in b:
        b.append(bb[i])
c={}#各分表数据字典
for i in range(len(b)):
    k=0
    c[b[i]]=[a[0]]
    for j in range(len(a)):
        if b[i]==a[j][ll]:  #按第ll列拆分
            c[b[i]] +=[a[j]]
for i in list(c.keys()):
    f=Workbook()
    sheet1=f.add_sheet('1')
    for j in range(len(c[i])):# j 行
        if len(c[i])>1:
            for k in range(col):#k 列
                sheet1.write(j,k,c[i][j][k])
    f.save(i.strip('\t')+'.xls')
print('已生成文件：',b)
'''