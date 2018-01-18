import xlrd
import xlwt
import glob
import csv
from tkinter import *
from tkinter.filedialog import askdirectory


def selectPath():
    path_ = askdirectory()
    way=path_+'/*.csv'
    path.set(path_)
    lst = glob.glob(way) 
    lst2=[]
    for list in lst:
       
    #bk = xlrd.open_workbook('D:/python/1.xlsx')
    #sh = bk.sheet_by_index(0)
        c=open(list,"r") #以r的方式打开csv文件
        read=csv.reader(c)
        test=list.replace('csv','xls')
        test2=test[len(way)-5:]
        #lst2.append(test)
        #print (lst)
        file_path = 'D:\\solarcell\\'+test2
        file = xlwt.Workbook()  #创建一个工作簿
        table = file.add_sheet('sheet 1')  #创建一个工作表
        n=0
        for line in read:
           
            if n>46:
                a=line[2]
                b=line[3]
                table.write(0,0,test2)
                table.write(n-47,1,a) #写入
                table.write(n-47,2,b) 
            n+=1
        file.save(file_path)  #保存'''  
root = Tk()
path = StringVar()
Label(root,text = "首先在D盘新建一个solarcell文件夹，单击路径选择后选择IV曲线存储的文件夹即可").grid(row = 0, column = 0,columnspan = 3)
Label(root,text = "目标路径:").grid(row = 1,column = 0)
Entry(root, textvariable = path).grid(row = 1, column = 1)
Button(root, text = "路径选择", command = selectPath).grid(row = 1, column = 2)
root.mainloop()

