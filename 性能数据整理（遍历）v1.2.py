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
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            list=os.path.join(parent,filename)
            if '.csv' in list and '-d' not in list and '-D' not in list and '-0' not in list and '_dark' not in list and '_d' not in list:
                c=open(list,"r") #以r的方式打开csv文件
                read=csv.reader(c)
                test=list.replace('csv','xls')
                test1=test[len(rootdir):]
                test2=test1.replace('.xls','')
                file_path = 'D:\\solarcell\\'+'汇总.xls'
                #file = xlwt.Workbook()  #创建一个工作簿
                #table = file.add_sheet('sheet 1')  #创建一个工作表
                style = xlwt.easyxf(num_format_str='#,##0.00')
                bk = xlrd.open_workbook('D:/solarcell/汇总.xls') #获取表格中已有行数
                sh = bk.sheet_by_index(0)
                n = 0
                i = sh.nrows
                bkcopy=copy(bk)
                shcopy=bkcopy.get_sheet(0)
                try:
                    for line in read:
                        if n == 1 and len(line)>10:
                            shcopy.write(i,0,test2,style)
                            for j in range(6,12):
                                a=line[j]
                                shcopy.write(i,j-5,a,style)
                        elif n==1 and len(line)<10:
                            shcopy.write(i,0,test2,style)
                        n+=1
                    bkcopy.save(file_path)  #保存'''  
                except:
                    print(test2)
root = Tk()
path = StringVar()
Label(root,text = "在D盘新建solarcell文件夹，文件夹下新建“汇总.xls”，将表格第一列重命名为“文件名  Jsc mA/cm2  Voc/V  FF  eff%  rsh  rs”单击路径选择后选择IV曲线存储的文件夹").grid(row = 0, column = 0, columnspan = 3)
Label(root,text = "目标路径:").grid(row = 2,column = 0)
Entry(root, textvariable = path).grid(row = 2, column = 1)
Button(root, text = "路径选择", command = selectPath).grid(row = 2, column = 2)
root.mainloop()

