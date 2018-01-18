import xlrd
import xlwt
import glob
import csv
from xlutils.copy import copy
from tkinter import *
from tkinter.filedialog import askdirectory
import os
import os.path

def selectPath():
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):  
        for filename in filenames:                        #输出文件信息
            list=os.path.join(parent,filename)
            #print(list)
            if '.csv' in list and '-d' not in list and '-D' not in list and '-0' not in list and '_dark' not in list and '_d' not in list:
        #bk = xlrd.open_workbook('D:/python/1.xlsx')
        #sh = bk.sheet_by_index(0)
                c=open(list,"r") #以r的方式打开csv文件
                read=csv.reader(c)
                test=list.replace('csv','xls')
                test1=test[len(rootdir):]
                test2=test1.replace('\\',' ')
                #lst2.append(test)
                #print (lst)
                file_path = 'D:\\solarcell\\'+test2
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
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            list1=os.path.join(parent,filename)
            list=list1.replace('/','\\')
            if '.xls' in list and '汇总' not in list:
                data=xlrd.open_workbook(list)
                sheet1=data.sheet_by_index(0)
                a=0
                if sheet1.nrows==600:
                    for i1 in range(600):
                        if float(sheet1.cell(i1,0).value)>0 and a == 0:
                            Isc1 = float(sheet1.cell(i1,1).value) - (float(sheet1.cell(i1,1).value) - float(sheet1.cell(i1-1,1).value))/(float(sheet1.cell(i1,0).value) - float(sheet1.cell(i1-1,0).value)) * float(sheet1.cell(i1,0).value)
                            Isc = Isc1 / 20.28*1000*(-1)
                            bkcopy=copy(data)
                            shcopy=bkcopy.get_sheet(0)
                            shcopy.write(1,6,Isc)
                            bkcopy.save(list)  #保存
                            a+=1
                    #print(Isc)
    #    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
    #        for filename in filenames:                        #输出文件信息
    #            list1=os.path.join(parent,filename)
    #            list=list1.replace('/','\\')
    #            if '.xls' in list and '汇总' not in list:
    #                data=xlrd.open_workbook(list)
    #                sheet1=data.sheet_by_index(0)
                    b=0
                    for i2 in range(600):
                        if float(sheet1.cell(i2,1).value)>0 and b == 0:
                            Voc = float(sheet1.cell(i2,0).value) - (float(sheet1.cell(i2,0).value) - float(sheet1.cell(i2-1,0).value))/(float(sheet1.cell(i2,1).value) - float(sheet1.cell(i2-1,1).value)) * float(sheet1.cell(i2,1).value) 
                            bkcopy=copy(data)
                            shcopy=bkcopy.get_sheet(0)
                            shcopy.write(1,7,Voc)
                            bkcopy.save(list)  #保存
                            b+=1
                    #print(Voc)
        #for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
            #for filename in filenames:                        #输出文件信息
                #list1=os.path.join(parent,filename)
                #list=list1.replace('/','\\')
                #if '.xls' in list and '汇总' not in list:
                    #data=xlrd.open_workbook(list)
                    #sheet1=data.sheet_by_index(0)
                    pmax=0
                    for j in range(600):
                        if float(sheet1.cell(j,0).value)*float(sheet1.cell(j,1).value)*(-1) > pmax:
                            pmax = float(sheet1.cell(j,0).value)*float(sheet1.cell(j,1).value)*(-1)
                    eff = pmax/20.28*1000
                    #print(Isc,Voc)
                    Pmax=Isc*Voc
                    FF = eff/Pmax
                    #style = xlwt.easyxf(num_format_str='#,##0.0000')
                    bkcopy=copy(data)
                    shcopy=bkcopy.get_sheet(0)
                    shcopy.write(1,8,FF)
                    shcopy.write(1,9,eff)
                    bkcopy.save(list)  #保存
        for filename in filenames:                        #输出文件信息
            list1=os.path.join(parent,filename)
            list=list1.replace('/','\\')
            if '.xls' in list and '汇总' not in list:
                data=xlrd.open_workbook(list)
                sheet1=data.sheet_by_index(0)
                a=0
                if sheet1.nrows==600:
                    for i1 in range(600):
                        if float(sheet1.cell(i1,0).value)>0 and a == 0:
                            Isc1 = float(sheet1.cell(i1,1).value) - (float(sheet1.cell(i1,1).value) - float(sheet1.cell(i1-1,1).value))/(float(sheet1.cell(i1,0).value) - float(sheet1.cell(i1-1,0).value)) * float(sheet1.cell(i1,0).value)
                            Isc = Isc1 / 20.28*1000*(-1)
                            bkcopy=copy(data)
                            shcopy=bkcopy.get_sheet(0)
                            shcopy.write(1,6,Isc)
                            bkcopy.save(list)  #保存
                            a+=1
        for filename in filenames:                        #输出文件信息
            list1=os.path.join(parent,filename)
            list=list1.replace('/','\\')
            if '.xls' in list and '汇总' not in list:
                data=xlrd.open_workbook(list)
                sheet1=data.sheet_by_index(0)
                if sheet1.nrows==600:
                    b=0
                    for i2 in range(600):
                        if float(sheet1.cell(i2,1).value)>0 and b == 0:
                            Voc = float(sheet1.cell(i2,0).value) - (float(sheet1.cell(i2,0).value) - float(sheet1.cell(i2-1,0).value))/(float(sheet1.cell(i2,1).value) - float(sheet1.cell(i2-1,1).value)) * float(sheet1.cell(i2,1).value) 
                            bkcopy=copy(data)
                            shcopy=bkcopy.get_sheet(0)
                            shcopy.write(1,7,Voc)
                            bkcopy.save(list)  #保存
                            b+=1                                                        
def addall():
    rootdir = askdirectory()
    path.set(rootdir)
    file_path = 'D:\\'+'汇总.xls'
    bk1 = xlrd.open_workbook('D:/汇总.xls') #获取表格中已有行数
    sh1 = bk1.sheet_by_index(0)
    k = sh1.nrows
    bkcopy1=copy(bk1)
    shcopy1=bkcopy1.get_sheet(0) 
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            list1=os.path.join(parent,filename)
            list=list1.replace('/','\\')
            test=list.replace('.xls','')
            test1=test[len(rootdir)+1:]
            if '.xls' in list and '汇总' not in list:
                data=xlrd.open_workbook(list)
                sheet1=data.sheet_by_index(0)
                shcopy1.write(k,0,test1)
                try:
                    if sheet1.cell(1,6).value:
                        shcopy1.write(k,1,sheet1.cell(1,6).value)
                    if sheet1.cell(1,7).value:                    
                        shcopy1.write(k,2,sheet1.cell(1,7).value)
                    if sheet1.cell(1,8).value:                    
                        shcopy1.write(k,3,sheet1.cell(1,8).value)
                    if sheet1.cell(1,9).value:                    
                        shcopy1.write(k,4,sheet1.cell(1,9).value)
                        bkcopy1.save(file_path)
                    if float(sheet1.cell(599,0).value) > 1:
                        shcopy1.write(k,7,'偏压超过1V')
                        bkcopy1.save(file_path)
                    k+=1#保存
                except:
                    print(list)
root = Tk()
path = StringVar()
Label(root,text = "首先在D盘新建一个solarcell文件夹，在D盘新建一个汇总xls，单击路径选择后选择IV曲线存储的文件夹即可").grid(row = 0, column = 0,columnspan = 5)
Label(root,text = "目标路径:").grid(row = 1,column = 0)
Entry(root, textvariable = path).grid(row = 1, column = 1)
Button(root, text = "路径选择", command = selectPath).grid(row = 1, column = 2)
Button(root, text = "计算结果", command = calculate).grid(row = 1, column = 3)
Button(root, text = "统计结果", command = addall).grid(row = 1, column = 4)

root.mainloop()

