from xlrd import *
from xlwt import *
from tkinter import *
from tkinter import filedialog
def chaifen(file,sheeti,colu):#拆分表格
    file=file+".xls"
    data = open_workbook(file)

    table3 =data.sheet_by_index(sheeti-1)
    a = []  # 所有数据
    colu = colu - 1  # 按第几列拆分 0
    # kk=input('请输入分列依据:')
    for i in range(table3.nrows):
        a += [table3.row_values(i)]
    bb = list(table3.col_values(colu))  # 按第ll列拆分
    b = []  # 分表字典键值
    col = len(table3.row_values(0))
    print("总列数：",col)
    for i in range(1, len(bb)):
        if bb[i] not in b:
            b.append(bb[i])
    c = {}  # 各分表数据字典
    bbb="已生成文件:共"+str(len(b))+'个\n'
    for i in range(len(b)):
        k = 0
        if i ==len(b)-1:
            bbb+=str(b[i])+'！'
        else:bbb+=str(b[i])+','
        c[b[i]] = [a[0]]
        for j in range(len(a)):
            if b[i] == a[j][colu]:  # 按第ll列拆分
                c[b[i]] += [a[j]]
    for i in list(c.keys()):
        f = Workbook()
        sheet1 = f.add_sheet('1')
        for j in range(len(c[i])):  # j 行
            if len(c[i]) > 1:
                for k in range(col):  # k 列
                    sheet1.write(j, k, c[i][j][k])
        f.save(i.strip('\t') + '.xls')
    print('已生成文件：', b)
    root1 = Tk(className='处理结果')
    textLabel = Label(root1, justify=LEFT)
    var1 = "文件名:<< % s>>" % e1.get()
    var2 = "第几张表:<< % s>>" % e2.get()
    var3 = "按第几列拆分:<< % s>>" % e2.get()
    textLabel['text'] = var1 + '\n' + var2 + '\n' + var3 + "\n"+bbb
    textLabel.pack(padx=10, pady=10,wraplength=20)
    theButton = Button(root1, text="确定", width=10, command=root.quit)
    theButton.pack(padx=10, pady=5)
def show(): #执行
    fil = e1.get()
    sht = int(e2.get())
    ll = int (e3.get())
    chaifen(fil, sht, ll)
    #e1.delete(0, END)
    #e2.delete(0, END)
    #e3.delete(0, END)
    #root.quit()

# 如果表格大于组件，那么可以使用sticky选项来设置组件的位置
# 同样你需要使用N，E，S,W以及他们的组合NE，SE，SW，NW来表示方位

root = Tk(className="拆分表格")
# Thinker总共提供了三种布局组件的方法：pack(),grid()和place()
# grid()方法允许你用表格的形式来管理组件的位置
# row选项代表行，column选项代表列
# 例如row=1，column=2表示第二行第三列(0表示第一行)
Label(root, text="文件名:").grid(row=0)
Label(root, text="第几张表:").grid(row=1)
Label(root, text="按第几列拆分:").grid(row=2)
e1 = Entry(root)
e2 = Entry(root)
e3 = Entry(root)
e1.grid(row=0, column=1, padx=10, pady=5)
e2.grid(row=1, column=1, padx=10, pady=5)
e3.grid(row=2, column=1, padx=10, pady=5)
Button(root, text="开始拆分", width=10, command=show).grid(row=3, column=0, sticky=W, padx=10, pady=5)
Button(root, text="退出", width=10, command=root.quit).grid(row=3, column=1, sticky=E, padx=10, pady=5)
mainloop()