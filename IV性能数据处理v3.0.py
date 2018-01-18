import xlrd
import xlwt
import csv
from tkinter import *
from tkinter.filedialog import askdirectory
import os
from xlutils.copy import copy
#先将csv数据中IV列写入xls中，后对xls进行处理并汇总

def simpleadd():
    global Dark
    rootdir = askdirectory()
    path.set(rootdir)
    try:
        file = xlwt.Workbook()  #创建一个工作簿
        table = file.add_sheet('sheet 1')  #创建一个工作表
        table.write(0,0,'文件路径')
        table.write(0,1,'Jsc mA/cm2')
        table.write(0,2,'Voc V')
        table.write(0,3,'FF')
        table.write(0,4,'Eff %')
        table.write(0,5,'rsh')
        table.write(0,6,'rs')
        file.save('D:/汇总.xls')
    except:
        print('汇总文件已存在')
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        
            list=os.path.join(parent,filename)
            judge = 1
            for v in Dark:
                if v in filename:
                    judge -=1
            if '.csv' in filename and judge ==1:
                c=open(list,"r") #以r的方式打开csv文件
                read=csv.reader(c)                
                path_=list[len(rootdir)+1:]                
                file_path = 'D:/汇总.xls'
                style = xlwt.easyxf(num_format_str='#,##0.00')
                bk = xlrd.open_workbook('D:/汇总.xls') #获取表格中已有行数
                sh = bk.sheet_by_index(0)
                n = 0
                i = sh.nrows
                bkcopy=copy(bk)
                shcopy=bkcopy.get_sheet(0)
                try:
                    for line in read:
                        if n == 1 and len(line)>10:
                            shcopy.write(i,0,path_,style)
                            for j in range(6,12):
                                a=line[j]
                                shcopy.write(i,j-5,a,style)
                        elif n==1 and len(line)<10:
                            shcopy.write(i,0,path_,style)
                        n+=1
                    bkcopy.save(file_path)  #保存
                except:
                    print(path_)
                    
def selectPath():
    global rootdir
    global Dark
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):  #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        
            list=os.path.join(parent,filename)
            judge = 1
            for v in Dark:
                if v in filename:
                    judge -=1
            if '.csv' in filename and judge ==1:
                c=open(list,"r") #以r的方式打开csv文件
                read=csv.reader(c)
                file_path = list.replace('csv','xls')
                file = xlwt.Workbook()  #创建一个工作簿
                table = file.add_sheet('sheet 1')  #创建一个工作表
                n=0
                for line in read:
                    if n>46 and len(line)>3:
                        a=line[2]
                        b=line[3]
                        table.write(n-47,0,a) #写入
                        table.write(n-47,1,b) 
                    n+=1
                file.save(file_path)  #保存

def calculate():
    global rootdir
    global Dark
    
    
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        
            list=os.path.join(parent,filename)
            if '.xls' in list and '汇总' not in list:
                data=xlrd.open_workbook(list)
                sheet1=data.sheet_by_index(0)
                if sheet1.nrows==600:
                    a=0
                    for i1 in range(600):
                        if float(sheet1.cell(i1,0).value)>0 and a == 0:
                            Isc1 = float(sheet1.cell(i1,1).value) - (float(sheet1.cell(i1,1).value) - float(sheet1.cell(i1-1,1).value))/(float(sheet1.cell(i1,0).value) - float(sheet1.cell(i1-1,0).value)) * float(sheet1.cell(i1,0).value)
                            Isc = Isc1 / 20.28*1000*(-1)
                            bkcopy=copy(data)
                            shcopy=bkcopy.get_sheet(0)
                            a+=1
                    b=0
                    for i2 in range(600):
                        if float(sheet1.cell(i2,1).value)>0 and b == 0:
                            Voc = float(sheet1.cell(i2,0).value) - (float(sheet1.cell(i2,0).value) - float(sheet1.cell(i2-1,0).value))/(float(sheet1.cell(i2,1).value) - float(sheet1.cell(i2-1,1).value)) * float(sheet1.cell(i2,1).value) 
                            bkcopy=copy(data)
                            shcopy=bkcopy.get_sheet(0)
                            b+=1
                    pmax=0
                    for j in range(600):
                        if float(sheet1.cell(j,0).value)*float(sheet1.cell(j,1).value)*(-1) > pmax:
                            pmax = float(sheet1.cell(j,0).value)*float(sheet1.cell(j,1).value)*(-1)
                    eff = pmax/20.28*1000
                    Pmax=Isc*Voc
                    FF = eff/Pmax
                    bkcopy=copy(data)
                    shcopy=bkcopy.get_sheet(0)
                    shcopy.write(1,6,Isc)
                    shcopy.write(1,7,Voc)
                    shcopy.write(1,8,FF)
                    shcopy.write(1,9,eff)
                    bkcopy.save(list)  #保存
                                 
def addall():
    global rootdir

    file = xlwt.Workbook()  #创建一个工作簿
    table = file.add_sheet('sheet 1')  #创建一个工作表
    file_path = rootdir + '/IV数据汇总.xls'    
    table.write(0,0,'测试文件路径')
    table.write(0,1,'Isc')
    table.write(0,2,'Voc')
    table.write(0,3,'FF')
    table.write(0,4,'Eff')
    k = 1
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        
            list1=os.path.join(parent,filename)
            list=list1.replace('\\','/')
            test=list[len(rootdir)+1:]
            if '.xls' in list and '汇总' not in list:
                data=xlrd.open_workbook(list)
                sheet1=data.sheet_by_index(0)
                table.write(k,0,test)
                try:
                    if sheet1.cell(1,6).value:
                        table.write(k,1,sheet1.cell(1,6).value)
                    if sheet1.cell(1,7).value:                    
                        table.write(k,2,sheet1.cell(1,7).value)
                    if sheet1.cell(1,8).value:                    
                        table.write(k,3,sheet1.cell(1,8).value)
                    if sheet1.cell(1,9).value:                    
                        table.write(k,4,sheet1.cell(1,9).value)
                    if float(sheet1.cell(599,0).value) > 1:
                        table.write(k,5,'施加偏压超过1V,为'+str(sheet1.cell(599,0).value)+'V')
                except:
                    print(list)
                k+=1#保存
    file.save(file_path)

def getdark():
    global Dark
    if dark.get():
        s = dark.get().split()
        Dark+=s
    else:
        print("非法输入！")

def simple():
    selectPath()
    calculate()
    addall()

rootdir = ""
Dark = ["dark","DARK"]

root = Tk()
root.title('IV性能数据处理 v3.0')
dark = StringVar()
path = StringVar()
Label(root,text = "单击路径选择后选择IV曲线存储的文件夹,后依次选择计算结果，统计结果").grid(row = 0, column = 0,columnspan = 6)
Label(root,text = "目标路径:").grid(row = 1,column = 0)
Entry(root, textvariable = path).grid(row = 1, column = 1,columnspan = 3)
Button(root, text = "计算结果", command = simple).grid(row = 1, column = 4)
Button(root, text = "读取已计算结果", command = simpleadd).grid(row = 1, column = 5)
Label(root,text = "请输入暗电流曲线包含的特征名称（空格分隔），默认为dark,DARK").grid(row = 2, column = 0,columnspan = 3)
Entry(root, textvariable = dark).grid(row = 2, column = 3,columnspan = 2)
Button(root, text = "输入名称", command = getdark).grid(row = 2, column = 5)
root.mainloop()

