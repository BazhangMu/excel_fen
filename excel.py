import pandas as pd
import xlrd
import xlwt
wb=pd.read_excel('D:\python3\excel_test\整改台账.xls',sheet_name='1')
wb1=pd.read_excel('D:\python3\excel_test\整改台账.xls',sheet_name='2')
jg=wb["机构"]
print(wb["机构"][0])
dic=
'''
#print(jg)
#print(type(jg))
jjg=[]
print('1:',len(jg))
#print('2:', jg.all)
for r in range(len(jg)):
    print(jg[r])
next()
print(jjg)
wb1.merge(wb[['检查内容','具体问题','存在问题网点']],on='整改情况')
#print(wb1)
#print(wb[['检查内容','具体问题']])
#print(pd.merge(wb, wb1))
wb3=pd.merge(wb1, wb)

#print(wb3)
#print(wb3['具体问题'])
#检查内容 	具体问题	存在问题网点	责任人	整改情况	机构
'''


