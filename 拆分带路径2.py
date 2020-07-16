from xlrd import *
from xlwt import *
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
global file_name,file_path,pathlab,root
file_name=''
file_path=''
def chaifen(file,sheeti,colu,path):#拆分表格
    global ret,root
    if file=='':
        file=e1.get()
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
    continuecf(len(b))
    if ret ==1 :
        bbb = ''
        for i in range(len(b)):
            k = 0
            if i == len(b) - 1:
                bbb += str(i+1)+str(b[i]).strip('\t') + '！'
            else:
                bbb += str(i+1)+str(b[i]).strip('\t') + '、'
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
            if path != '':
                f.save(path + "/" + str(i).strip('\t') + '.xls')
            else:
                f.save(str(i).strip('\t') + '.xls')
        print('已生成文件：', b)
        var1 = "文件名:<< % s>>" % e1.get()
        var2 = "第几张表:<< % s>>" % e2.get()
        var3 = "按第几列拆分:<< % s>>" % e3.get()
        var4 = "新文件路径:<< % s>>" % path
        messagebox.showinfo(title='处理结果', message=var1 + '\n' + var2 + '\n' + var3 + "\n" + var4 + "\n" +"已生成文件: 共" + str(len(b)) + "个\n"+bbb)
        rootc.destroy()
    else:
        var1 = "文件名:<< % s>>" % e1.get()
        var2 = "第几张表:<< % s>>" % e2.get()
        var3 = "按第几列拆分:<< % s>>" % e3.get()
        var4 = "结果-----> 未执行拆分！"
        messagebox.showinfo(title='未处理', message=var1 + '\n' + var2 + '\n' + var3 + "\n" + var4)
        rootc.destroy()
def continuecf(numb):
    global rootc
    rootc=Tk(className="是否继续?")
    Label(rootc, text="即将执行拆分，预计生成文件个数: "+str(numb)).grid(row=0)
    Button(rootc, text="继续", width=10, command=continuey).grid(row=1, column=0, sticky=W,padx=10, pady=5)
    Button(rootc, text="退出", width=10, command=continuen).grid(row=1, column=1, sticky=E, padx=10, pady=5)
    mainloop()
def continuey():
    global ret,rootc
    ret=1
    rootc.quit()
def continuen():
    global ret,rootc
    ret=0
    rootc.quit()
def show(): #执行
    global root
    fil = e1.get()
    sht = int(e2.get())
    ll = int (e3.get())
    chaifen(file_name, sht, ll,file_path)
def chosefile():#选择文件
    global file_path,file_name,pathlab
    rootc = Tk()
    rootc.withdraw()
    file_name = filedialog.askopenfilename()  # 文件名
    e1padx=len(file_name)+10
    #print(e1padx)
    e1.grid(row=0, column=1, ipadx=e1padx)  #-column列, -columnspan偏移, -in, -ipadx文本宽, -ipady, -padx, -pady, -row, -rowspan, or -stickye1pady*1.5
    e2.grid(row=1, column=1, padx=10, ipadx=e1padx)
    e3.grid(row=2, column=1, padx=10, ipadx=e1padx)
    e2.delete(0, END)
    e3.delete(0, END)
    #print(file_name)
    if file_name !='':
        e1.delete(0, END)
        Entry.insert(e1, 0, file_name)
        fullpath = file_name.split("/")
        file_path = ''
        for i in range(len(fullpath) - 1):
            file_path += fullpath[i] + "/"
        pathlab.grid_remove()
        pathlab = Label(root, text=file_path)
        pathlab.grid(row=3, column=1, sticky=W)
    else:
        e1.delete(0, END)
        Entry.insert(e1, 0, "请录入程序所在目录下文件名")
def chosepath():#选择路径
    global file_path,pathlab
    rootc = Tk()
    rootc.withdraw()
    file_path = filedialog.askdirectory()  # 文件夹
    print(type(file_path))
    if file_path != '':
        pathlab.grid_remove()
        pathlab=Label(root, text=file_path)
        pathlab.grid(row=3, column=1, sticky=W)
    else:
        pathlab.grid_remove()
        pathlab =Label(root, text = "原文件路径")
        pathlab.grid(row=3, column=1, sticky=W)
root = Tk(className="拆分表格")
Label(root, text="文件名:").grid(row=0)
Label(root, text="第几张表:").grid(row=1)
Label(root, text="按第几列拆分:").grid(row=2)
Label(root, text="输出路径:").grid(row=3)
pathlab=Label(root, text="原文件路径")
pathlab.grid(row=3, column=1, sticky=W)
e1 = Entry(root)
Entry.insert(e1, 0, '可手工录入' )
e2 = Entry(root)
Entry.insert(e2, 0, '输入拆分表序号' )
e3 = Entry(root)
Entry.insert(e3, 0, '输入拆分列序号' )
e1.grid(row=0, column=1, padx=10, pady=5)
e2.grid(row=1, column=1, padx=10, pady=5)
e3.grid(row=2, column=1, padx=10, pady=5)
Button(root, text="选择文件", width=10, command=chosefile).grid(row=0, column=2,  padx=10, pady=5)
Button(root, text="开始拆分", width=10, command=show).grid(row=4, column=0, sticky=W, padx=10, pady=5)
Button(root, text="退出", width=10,command=root.quit).grid(row=4, column=1, sticky=E, padx=10, pady=5)
Button(root, text="选择路径", width=10, command=chosepath).grid(row=3, column=2,  padx=10, pady=5)
mainloop()