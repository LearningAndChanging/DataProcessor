from tkinter import *
from tkinter.filedialog import askdirectory
import os.path
import xml.etree.ElementTree as ET
import xlrd
import xlwt
import os
from xlutils.copy import copy

def changetoxls():
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            nroot=os.path.join(parent,filename)
            if '.xml' in nroot:
                fullpath = nroot.replace('\\','/')         #得到文件完整路径
                fullpathxls = nroot.replace('.xml','.xls')
                tree=ET.parse(fullpath)
                root1 = tree.getroot()
                n=0
                file = xlwt.Workbook()  #创建一个工作簿
                table = file.add_sheet('sheet 1')  #创建一个工作表
                table.write(0,1,'GP in 1/s')
                table.write(0,2,'Tau in Pa')
                table.write(0,3,'Eta in mPas')
                table.write(0,4,'T in oC')
                table.write(0,5,'t in s')
                table.write(0,6,'t_seg in s')   #写入表头
                if 'mPas'in root1[3][0][0][3][0].text:
                    changeunit = 1
                else:
                    changeunit = 1000
                for data in root1.iter(root1[3][0][0][0][0].tag):
                    if n>6 and (n%7)==3:
                        table.write(n//7,n%7,float(data.text)*changeunit)  #写入
                    elif n>6 and (n%7)!=0 and (n%7)!=3:
                        table.write(n//7,n%7,float(data.text))  #写入
                    elif n>6:
                        table.write(n//7,n%7,data.text)
                    n+=1
                file.save(fullpathxls)  #保存
                
def changetotxt():
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            nroot=os.path.join(parent,filename)
            if '.xml' in nroot:
                fullpath = nroot.replace('\\','/')         #得到文件完整路径  
                newname = fullpath.replace('.xml','.txt')
                os.rename(fullpath,newname)

def changetoxml():
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            nroot=os.path.join(parent,filename)
            if '.txt' in nroot:
                fullpath = nroot.replace('\\','/')         #得到文件完整路径  
                newname = fullpath.replace('.txt','.xml')
                os.rename(fullpath,newname)

def changetoutf8():
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            nroot=os.path.join(parent,filename)
            if '.txt' in nroot:
                fullpath = nroot.replace('\\','/')         #得到文件完整路径  
                f = open (fullpath, "r",encoding = 'utf-8',errors = 'ignore')
                con = f.read()
                newpath = fullpath.replace('.txt','-new.txt')
                open(fullpath, 'w+',encoding = 'utf-8').write(re.sub(r'xml version="1.0"', r'xml version="1.0"  encoding="UTF-8"', con))

def calculate():
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            nroot1=os.path.join(parent,filename)
            nroot=nroot1.replace('/','\\')
            if '.xls' in nroot and '汇总' not in nroot:
                data=xlrd.open_workbook(nroot)
                sheet1=data.sheet_by_index(0)
                
                if sheet1.nrows > 300:
                    EtaMax = 0
                    for j in range(1,301):
                        if float(sheet1.cell(j,3).value) > EtaMax:
                            EtaMax = float(sheet1.cell(j,3).value)
                            GPMax = float(sheet1.cell(j,1).value)
                    a = 0
                    for i1 in range(2,301):
                        if float(sheet1.cell(i1,1).value) > 10 and a == 0:
                            GP10 = float(sheet1.cell(i1,3).value) - (float(sheet1.cell(i1,3).value) - float(sheet1.cell(i1-1,3).value))/(float(sheet1.cell(i1,1).value) - float(sheet1.cell(i1-1,1).value)) * (float(sheet1.cell(i1,1).value)-10)
                            a+=1
                            
                    b=0
                    for i2 in range(2,301):
                        if float(sheet1.cell(i2,1).value) > 5 and b == 0:
                            GP5 = float(sheet1.cell(i2,3).value) - (float(sheet1.cell(i2,3).value) - float(sheet1.cell(i2-1,3).value))/(float(sheet1.cell(i2,1).value) - float(sheet1.cell(i2-1,1).value)) * (float(sheet1.cell(i2,1).value)-5)
                            b+=1
                    Eta100 = float(sheet1.cell(100,3).value)
                    Eta200 = float(sheet1.cell(200,3).value)
                    Eta300 = float(sheet1.cell(300,3).value)
                    bkcopy=copy(data)
                    shcopy=bkcopy.get_sheet(0)
                    
                    if EtaMax:
                        shcopy.write(1,8,EtaMax)
                    if GPMax:
                        shcopy.write(1,9,GPMax)
                    if GP5:
                        shcopy.write(1,10,GP5)
                    if GP10:
                        shcopy.write(1,11,GP10)
                    if Eta100:
                        shcopy.write(1,12,Eta100)
                    if Eta200:
                        shcopy.write(1,13,Eta200)
                    if Eta300:
                        shcopy.write(1,14,Eta300)
                    bkcopy.save(nroot)  #保存
                    
                                                   
def addall():
    rootdir = askdirectory()
    path.set(rootdir)
    file_path = 'D:\\'+'流变汇总.xls'
    bk1 = xlrd.open_workbook('D:/流变汇总.xls') #获取表格中已有行数
    sh1 = bk1.sheet_by_index(0)
    k = sh1.nrows
    bkcopy1=copy(bk1)
    shcopy1=bkcopy1.get_sheet(0) 
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            nroot1=os.path.join(parent,filename)
            nroot=nroot1.replace('/','\\')
            test=nroot.replace('.xls','')
            test1=test[len(rootdir)+1:]
            if '.xls' in nroot and '汇总' not in nroot:
                data=xlrd.open_workbook(nroot)
                sheet1=data.sheet_by_index(0)
                shcopy1.write(k,0,test1)
                try:
                    if sheet1.cell(1,8).value:
                        shcopy1.write(k,1,sheet1.cell(1,8).value)
                    if sheet1.cell(1,9).value:                    
                        shcopy1.write(k,2,sheet1.cell(1,9).value)
                    if sheet1.cell(1,10).value:                    
                        shcopy1.write(k,3,sheet1.cell(1,10).value)
                    if sheet1.cell(1,11).value:                    
                        shcopy1.write(k,4,sheet1.cell(1,11).value)
                    if sheet1.cell(1,12).value:                    
                        shcopy1.write(k,5,sheet1.cell(1,12).value)
                    if sheet1.cell(1,13).value:                    
                        shcopy1.write(k,6,sheet1.cell(1,13).value)
                    if sheet1.cell(1,14).value:                    
                        shcopy1.write(k,7,sheet1.cell(1,14).value)                        
                        bkcopy1.save(file_path)
                    k+=1#保存
                except:
                    print(nroot)

def collectall():
    rootdir = askdirectory()
    path.set(rootdir)
    file_path = 'D:\\'+'原始流变汇总.xls'
    bk1 = xlrd.open_workbook('D:/原始流变汇总.xls') 
    sh1 = bk1.sheet_by_index(0)
    k = sh1.nrows                                         #获取表格中已有行数
    bkcopy1=copy(bk1)
    shcopy1=bkcopy1.get_sheet(0) 
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            nroot1=os.path.join(parent,filename)
            nroot=nroot1.replace('/','\\')
            test=nroot.replace('.xls','')
            test1=test[len(rootdir)+1:]
            if '.xls' in nroot and '汇总' not in nroot:
                data=xlrd.open_workbook(nroot)
                sheet1=data.sheet_by_index(0)
                rownum=sheet1.nrows
                colnum=sheet1.ncols
                try:
                    for i in range(1,301):
                        for j in range(7):
                            if sheet1.cell(i,j).value:
                                shcopy1.write(k,j+1,sheet1.cell(i,j).value)                      
                        shcopy1.write(k,0,test1)
                        k+=1#保存
                except:
                    print(nroot)
    bkcopy1.save(file_path)
root = Tk()
root.title('RheologyData Processer v1.0')
path = StringVar()
Label(root,text = "单击目标方法后选择流变数据存储的文件夹").grid(row = 0, column = 0, columnspan = 4)
Label(root,text = "目标路径:").grid(row = 2,column = 0)
Entry(root, textvariable = path).grid(row = 2, column = 1, columnspan = 3)
Button(root, text = "ChangeToTxt", command = changetotxt).grid(row = 3, column = 0)
Button(root, text = "ChangeToUtf8", command = changetoutf8).grid(row = 3, column = 1)
Button(root, text = "ChangeToXml", command = changetoxml).grid(row = 3, column = 2)
Button(root, text = "ChangeToXls", command = changetoxls).grid(row = 3, column = 3)
Button(root, text = "Calculate", command = calculate).grid(row = 4, column = 0, columnspan = 1)
Button(root, text = "Addall", command = addall).grid(row = 4, column = 1, columnspan = 1)
Button(root, text = "Collectall", command = collectall).grid(row = 4, column = 2, columnspan = 1)
root.mainloop()

'''
import os
from tkinter import *
from tkinter.filedialog import askdirectory
import os.path

def changetotxt():
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            nroot=os.path.join(parent,filename)
            if '.xml' in nroot:
                fullpath = nroot.replace('\\','/')         #得到文件完整路径  
                newname = fullpath.replace('.xml','.txt')
                os.rename(fullpath,newname)

def changetoxml():
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            nroot=os.path.join(parent,filename)
            if '.txt' in nroot:
                fullpath = nroot.replace('\\','/')         #得到文件完整路径  
                newname = fullpath.replace('.txt','.xml')
                os.rename(fullpath,newname)

def changetoutf8():
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        #输出文件信息
            nroot=os.path.join(parent,filename)
            if '.txt' in nroot:
                fullpath = nroot.replace('\\','/')         #得到文件完整路径  
                f = open (fullpath, "r",encoding = 'utf-8',errors = 'ignore')
                con = f.read()
                newpath = fullpath.replace('.txt','-new.txt')
                open(fullpath, 'w+',encoding = 'utf-8').write(re.sub(r'xml version="1.0"', r'xml version="1.0"  encoding="UTF-8"', con))

root = Tk()
path = StringVar()
Label(root,text = "单击目标方法后选择流变数据存储的文件夹").grid(row = 0, column = 0, columnspan = 3)
Label(root,text = "目标路径:").grid(row = 2,column = 0)
Entry(root, textvariable = path).grid(row = 2, column = 1, columnspan = 2)
Button(root, text = "ChangeToTxt", command = changetotxt).grid(row = 3, column = 0)
Button(root, text = "ChangeToUtf8", command = changetoutf8).grid(row = 3, column = 1)
Button(root, text = "ChangeToXml", command = changetoxml).grid(row = 3, column = 2)
root.mainloop()
'''


'''
import mysql.connector
import mysql.connector
from mysql.connector import errorcode

cnx = mysql.connector.connect(user='root', password='jx0217',
                              host='127.0.0.1',
                              database='yinjiang')

TABLES = {}
TABLES['employees'] = (
    "CREATE TABLE `employees` ("
    "  `emp_no` int(11) NOT NULL AUTO_INCREMENT,"
    "  `birth_date` date NOT NULL,"
    "  `first_name` varchar(14) NOT NULL,"
    "  `last_name` varchar(16) NOT NULL,"
    "  `gender` enum('M','F') NOT NULL,"
    "  `hire_date` date NOT NULL,"
    "  PRIMARY KEY (`emp_no`)"
    ") ENGINE=InnoDB")

TABLES['departments'] = (
    "CREATE TABLE `departments` ("
    "  `dept_no` char(4) NOT NULL,"
    "  `dept_name` varchar(40) NOT NULL,"
    "  PRIMARY KEY (`dept_no`), UNIQUE KEY `dept_name` (`dept_name`)"
    ") ENGINE=InnoDB")

TABLES['salaries'] = (
    "CREATE TABLE `salaries` ("
    "  `emp_no` int(11) NOT NULL,"
    "  `salary` int(11) NOT NULL,"
    "  `from_date` date NOT NULL,"
    "  `to_date` date NOT NULL,"
    "  PRIMARY KEY (`emp_no`,`from_date`), KEY `emp_no` (`emp_no`),"
    "  CONSTRAINT `salaries_ibfk_1` FOREIGN KEY (`emp_no`) "
    "     REFERENCES `employees` (`emp_no`) ON DELETE CASCADE"
    ") ENGINE=InnoDB")

TABLES['dept_emp'] = (
    "CREATE TABLE `dept_emp` ("
    "  `emp_no` int(11) NOT NULL,"
    "  `dept_no` char(4) NOT NULL,"
    "  `from_date` date NOT NULL,"
    "  `to_date` date NOT NULL,"
    "  PRIMARY KEY (`emp_no`,`dept_no`), KEY `emp_no` (`emp_no`),"
    "  KEY `dept_no` (`dept_no`),"
    "  CONSTRAINT `dept_emp_ibfk_1` FOREIGN KEY (`emp_no`) "
    "     REFERENCES `employees` (`emp_no`) ON DELETE CASCADE,"
    "  CONSTRAINT `dept_emp_ibfk_2` FOREIGN KEY (`dept_no`) "
    "     REFERENCES `departments` (`dept_no`) ON DELETE CASCADE"
    ") ENGINE=InnoDB")

TABLES['dept_manager'] = (
    "  CREATE TABLE `dept_manager` ("
    "  `dept_no` char(4) NOT NULL,"
    "  `emp_no` int(11) NOT NULL,"
    "  `from_date` date NOT NULL,"
    "  `to_date` date NOT NULL,"
    "  PRIMARY KEY (`emp_no`,`dept_no`),"
    "  KEY `emp_no` (`emp_no`),"
    "  KEY `dept_no` (`dept_no`),"
    "  CONSTRAINT `dept_manager_ibfk_1` FOREIGN KEY (`emp_no`) "
    "     REFERENCES `employees` (`emp_no`) ON DELETE CASCADE,"
    "  CONSTRAINT `dept_manager_ibfk_2` FOREIGN KEY (`dept_no`) "
    "     REFERENCES `departments` (`dept_no`) ON DELETE CASCADE"
    ") ENGINE=InnoDB")

TABLES['titles'] = (
    "CREATE TABLE `titles` ("
    "  `emp_no` int(11) NOT NULL,"
    "  `title` varchar(50) NOT NULL,"
    "  `from_date` date NOT NULL,"
    "  `to_date` date DEFAULT NULL,"
    "  PRIMARY KEY (`emp_no`,`title`,`from_date`), KEY `emp_no` (`emp_no`),"
    "  CONSTRAINT `titles_ibfk_1` FOREIGN KEY (`emp_no`)"
    "     REFERENCES `employees` (`emp_no`) ON DELETE CASCADE"
    ") ENGINE=InnoDB")

'''
'''
DB_NAME = 'test3'
def create_database(cursor):
    try:
        cursor.execute(
            "CREATE DATABASE {} DEFAULT CHARACTER SET 'utf8'".format(DB_NAME))
    except mysql.connector.Error as err:
        print("Failed creating database: {}".format(err))
        exit(1)
      
cursor = cnx.cursor()
try:
    cnx.database = DB_NAME  
except mysql.connector.Error as err:
    if err.errno == errorcode.ER_BAD_DB_ERROR:
        create_database(cursor)
        cnx.database = DB_NAME
    else:
        print(err)
        exit(1) #建立数据库
'''
'''
for name, ddl in TABLES.items():
    try:
        print("Creating table {}: ".format(name), end='')
        cursor.execute(ddl)
    except mysql.connector.Error as err:
        if err.errno == errorcode.ER_TABLE_EXISTS_ERROR:
            print("already exists.")
        else:
            print(err.msg)
    else:
        print("OK")

cursor.close()
cnx.close()
'''
