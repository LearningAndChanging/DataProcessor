import xlrd
import xlwt
import os
import numpy as np
import matplotlib
from tkinter import *
from tkinter import filedialog
from pylab import annotate
from xlutils.copy import copy
import csv
from tkinter.filedialog import askdirectory
import pickle
import hashlib
import pymysql

class CSVManager(object):
    def __init__(self):
        self.new_csvs = self.load_progress('new_csvs.txt')#未爬取URL集合
        self.old_csvs = self.load_progress('old_csvs.txt')#已爬取URL集合
    def has_new_csv(self):
        '''
        判断是否有未处理的csv
        :return:
        '''
        return self.new_csv_size()!=0

    def get_new_csv(self):
        '''
        获取一个未爬取的csv
        :return:
        '''
        new_csv = self.new_csvs.pop()
        m = hashlib.md5()
        csv_data = open(new_csv, 'rb')
        m.update(csv_data.read())
        csv_data.close()
        self.old_csvs.add(m.hexdigest())
        return new_csv

    def add_new_csv(self,csv):
        '''
         将新的csv添加到未处理的csv集合中
        :param csv:单个csv
        :return:
        '''
        if csv is None:
            return
        m = hashlib.md5()
        csv_data = open(csv, 'rb')
        m.update(csv_data.read())
        csv_data.close()
        csv_md5 =  m.hexdigest()
        if csv not in self.new_csvs and csv_md5 not in self.old_csvs:
            self.new_csvs.add(csv)
        return("ok")

    def add_new_csvs(self,csvs):
        '''
        将新的csvS添加到未处理的csv集合中
        :param csvs:csv集合
        :return:
        '''
        if csvs is None or len(csvs)==0:
            return
        for csv in csvs:
            self.add_new_csv(csv)

    def new_csv_size(self):
        '''
        获取未处理csv集合的大小
        :return:
        '''
        return len(self.new_csvs)

    def old_csv_size(self):
        '''
        获取已经处理csv集合的大小
        :return:
        '''
        return len(self.old_csvs)

    def save_progress(self,path,data):
        '''
        保存进度
        :param path:文件路径
        :param data:数据
        :return:
        '''
        with open(path, 'wb') as f:
            pickle.dump(data, f)

    def load_progress(self,path):
        '''
        从本地文件加载进度
        :param path:文件路径
        :return:返回set集合
        '''
        print ('[+] 从文件加载进度: %s' % path)
        try:
            with open(path, 'rb') as f:
                tmp = pickle.load(f)
                return tmp
        except:
            print ('[!] 无进度文件, 创建: %s' % path)
        return set()
            
def getarea():
    global area
    if Area.get() and float(Area.get()):
        area = float(Area.get())
    else:
        area = 27.04
        
def getvolt():
    global width
    if Width.get() and float(Width.get()):
        width = float(Width.get())
    else:
        width = 0.7
        
def getelec():
    global height
    if Height.get() and float(Height.get()):
        height = float(Height.get())
    else:
        height = 40

def getdark():
    global Dark
    if dark.get():
        s = dark.get().split()
        Dark+=s
    else:
        print("非法输入！")

def insert2db(tp):
    # 打开数据库连接
    db = pymysql.connect("localhost","root","root","yinjiang",charset='utf8')
    # 使用cursor()方法获取操作游标 
    cursor = db.cursor()
    # SQL 插入语句
    sql = "INSERT INTO IVData(DataRoot, \
           Isc, Voc, FF, Eff) \
           VALUES ('%s', '%f', '%f', '%f', '%f')" % \
           (tp[0], tp[1], tp[2], tp[3], tp[4])
    try:
       # 执行sql语句
       cursor.execute(sql)
       # 执行sql语句
       db.commit()
       print ("Data insert successfully!")
    except:
       # 发生错误时回滚
       db.rollback()
    # 关闭数据库连接
    db.close()
     
def targetselect():
    global area
    global width
    global height
    global Dark
    rootdir = filedialog.askopenfilename()
    path.set(rootdir)
    nowdir = rootdir
    drawonepng(nowdir,width,height)
    matplotlib.pyplot.show()

def documentselect():
    global area
    global width
    global height
    global Dark
    global docupath
    rootdir = askdirectory()
    docupath = rootdir.replace('\\','/')
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        
            list1=os.path.join(parent,filename)
            nowdir=list1.replace('\\','/')
            print(nowdir)
            judge = 1
            for v in Dark:
                if v in nowdir:
                    judge -=1
                    print("暗电流曲线")
            if '.csv' in nowdir and judge ==1:

                csvmanager=CSVManager()
                csvmanager.add_new_csv(nowdir)
                if csvmanager.has_new_csv():
                    csvmanager.get_new_csv()
                    print("开始读取")
                    data = drawonepng(nowdir)
                    matplotlib.pyplot.close()
                    if data:
                        insert2db(data)
                    csvmanager.save_progress('new_csvs.txt',csvmanager.new_csvs)
                    csvmanager.save_progress('old_csvs.txt',csvmanager.old_csvs)

def drawonepng(nowdir,width=0.7,height=40):
    global docupath
    global nfile_path
    global area
    global Dark
    judge = 1
    for v in Dark:
        if v in nowdir:
            judge -=1
            print("暗电流曲线")
    if '.csv' in nowdir and judge ==1:
        c=open(nowdir,"r") #以r的方式打开csv文件
        read=csv.reader(c)
        nfile_path = nowdir.replace(".csv",".xls")
        print("准备创建xls")
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
        file.save(nfile_path)  #保存
        print(nfile_path+"创建成功！")
    '''if '.xls' in nfile_path and '汇总' not in nfile_path:
        data=xlrd.open_workbook(nfile_path)
        sheet1=data.sheet_by_index(0)
        if sheet1.nrows==600:
            a=0
            for i1 in range(600):
                if float(sheet1.cell(i1,0).value)>0 and a == 0:
                    Isc1 = float(sheet1.cell(i1,1).value) - (float(sheet1.cell(i1,1).value) - float(sheet1.cell(i1-1,1).value))/(float(sheet1.cell(i1,0).value) - float(sheet1.cell(i1-1,0).value)) * float(sheet1.cell(i1,0).value)
                    Isc = Isc1 / area*1000*(-1)
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
            eff = pmax/area*1000
            Pmax=Isc*Voc
            FF = eff/Pmax
            bkcopy=copy(data)
            shcopy=bkcopy.get_sheet(0)
            shcopy.write(1,6,Isc)
            shcopy.write(1,7,Voc)
            shcopy.write(1,8,FF)
            shcopy.write(1,9,eff)
            bkcopy.save(nfile_path)  #保存
            print("Isc=",format(Isc,".2f"),"mA/cm2 Voc=",format(Voc,".2f"),"V FF=",format(FF*100,".2f"),"% Eff=",format(eff,".2f"),"%")
            if Isc<0 or Isc>50 or Voc<0 or Voc>0.7 or FF<0 or FF>0.9 or eff<5 or eff>20:
                return
    data = xlrd.open_workbook(nfile_path)
    sheet1 = data.sheet_by_index(0)
    x = []
    y = []
    for i in range(600):
        x.append(float(sheet1.cell(i,0).value))
        y.append(-float(sheet1.cell(i,1).value)/area*1000)
    # 通过rcParams设置全局横纵轴字体大小
    matplotlib.pyplot.rcParams['font.sans-serif']=['Arial']
    matplotlib.rcParams['xtick.direction'] = 'in' 
    matplotlib.rcParams['ytick.direction'] = 'in' 
    matplotlib.rcParams['xtick.labelsize'] = 12
    matplotlib.rcParams['ytick.labelsize'] = 12
    font = {'family' : 'Arial',
        'color'  : 'black',
        'weight' : 'normal',
        'size'   : 16,
        }
    
    matplotlib.pyplot.figure('IV Curve',figsize=(6,4.5))
    ax = matplotlib.pyplot.subplot(1,1,1)
    ax.spines['bottom'].set_linewidth(1.3)
    ax.spines['left'].set_linewidth(1.3)
    ax.spines['top'].set_linewidth(1.3)        
    ax.spines['right'].set_linewidth(1.3)
    # 通过'k'指定线的颜色，lw指定线的宽度
    # 第三个参数除了颜色也可以指定线形，比如'r--'表示红色虚线
    # 更多属性可以参考官网：http://matplotlib.org/api/pyplot_api.html
    matplotlib.pyplot.plot(x, y, 'r', lw=1.5)
    matplotlib.pyplot.xlim(0, width)
    matplotlib.pyplot.xlabel('U (V)',fontdict=font)
    matplotlib.pyplot.ylim(0, height)
    matplotlib.pyplot.ylabel('I (mA/cm2)',fontdict=font)
    # scatter可以更容易地生成散点图
    #matplotlib.pyplot.scatter(x, y)
    annotate('area  = ' + str(format(area,".2f")+' cm2'),
             xy=(0, 0), xycoords='data',
             xytext=(+10, +110), textcoords='offset points', fontsize=12)    
    annotate('Isc  = ' + str(format(Isc,".2f")+' mA/cm2'),
             xy=(0, 0), xycoords='data',
             xytext=(+10, +85), textcoords='offset points', fontsize=12)
    annotate('Voc = ' + str(format(Voc,".2f")+'  V'),
             xy=(0, 0), xycoords='data',
             xytext=(+10, +60), textcoords='offset points', fontsize=12)
    annotate('FF   = ' + str(format(FF*100,".2f")+' %'),
             xy=(0, 0), xycoords='data',
             xytext=(+10, +35), textcoords='offset points', fontsize=12)
    annotate('Eff  = ' + str(format(eff,".2f")+'  %'),
             xy=(0, 0), xycoords='data',
             xytext=(+10, +10), textcoords='offset points', fontsize=12)
    matplotlib.pyplot.grid(False)
    # 将当前figure的图保存到文件
    # 将结果汇总
    if docupath:
        xlsfile_path = docupath+'/汇总.xls'
        if not os.path.exists(xlsfile_path):
            file = xlwt.Workbook()  #创建一个工作簿
            table = file.add_sheet('sheet 1')  #创建一个工作表
            table.write(0,0,'文件路径')
            table.write(0,1,'Jsc mA/cm2')
            table.write(0,2,'Voc V')
            table.write(0,3,'FF')
            table.write(0,4,'Eff %')
            table.write(0,5,'rsh')
            table.write(0,6,'rs')
            file.save(xlsfile_path)
            print('汇总文件'+xlsfile_path+'创建成功!') 
        else:
            print('汇总文件已存在')    
        
        bkall = xlrd.open_workbook(xlsfile_path) #获取表格中已有行数
        nxlsfile_path = xlsfile_path.replace(".xls",'')
        shall = bkall.sheet_by_index(0)
        irow = shall.nrows
        print("汇总表格中已有"+str(irow)+"行")
        bkcopyall=copy(bkall)
        shcopyall=bkcopyall.get_sheet(0)
        try:
            #style = xlwt.easyxf(num_format_str='#,##0.00')
            shcopyall.write(irow,0,nowdir)
            shcopyall.write(irow,1,Isc)
            shcopyall.write(irow,2,Voc)
            shcopyall.write(irow,3,FF)
            shcopyall.write(irow,4,eff)
            bkcopyall.save(xlsfile_path)  #保存
        except:
            print(nowdir+"汇总失败！")
    # 删除xls文档
    os.remove(nfile_path)
    print("xls文件删除成功！")
    
    nnfile_path = nowdir.replace(".csv",".png")
    print(nnfile_path+"图片保存成功！")
    matplotlib.pyplot.savefig(nnfile_path, bbox_inches='tight', dpi=150)
    return (nowdir,Isc,Voc,FF,eff)'''
    

    try:
        if '.xls' in nfile_path and '汇总' not in nfile_path:
            data=xlrd.open_workbook(nfile_path)
            sheet1=data.sheet_by_index(0)
        
            if sheet1.nrows==600:
                a=0
                for i1 in range(600):
                    if float(sheet1.cell(i1,0).value)>0 and a == 0:
                        Isc1 = float(sheet1.cell(i1,1).value) - (float(sheet1.cell(i1,1).value) - float(sheet1.cell(i1-1,1).value))/(float(sheet1.cell(i1,0).value) - float(sheet1.cell(i1-1,0).value)) * float(sheet1.cell(i1,0).value)
                        Isc = Isc1 / area*1000*(-1)
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
                eff = pmax/area*1000
                Pmax=Isc*Voc
                FF = eff/Pmax
                bkcopy=copy(data)
                shcopy=bkcopy.get_sheet(0)
                shcopy.write(1,6,Isc)
                shcopy.write(1,7,Voc)
                shcopy.write(1,8,FF)
                shcopy.write(1,9,eff)
                bkcopy.save(nfile_path)  #保存
                print("Isc=",format(Isc,".2f"),"mA/cm2 Voc=",format(Voc,".2f"),"V FF=",format(FF*100,".2f"),"% Eff=",format(eff,".2f"),"%")
                if Isc<0 or Isc>40 or Voc<0 or Voc>0.7 or FF<0 or FF>0.9 or eff<5 or eff>20:
                    return
        data = xlrd.open_workbook(nfile_path)
        sheet1 = data.sheet_by_index(0)
        x = []
        y = []
        for i in range(600):
            x.append(float(sheet1.cell(i,0).value))
            y.append(-float(sheet1.cell(i,1).value)/area*1000)
        # 通过rcParams设置全局横纵轴字体大小
        matplotlib.pyplot.rcParams['font.sans-serif']=['Arial']
        matplotlib.rcParams['xtick.direction'] = 'in' 
        matplotlib.rcParams['ytick.direction'] = 'in' 
        matplotlib.rcParams['xtick.labelsize'] = 12
        matplotlib.rcParams['ytick.labelsize'] = 12
        font = {'family' : 'Arial',
            'color'  : 'black',
            'weight' : 'normal',
            'size'   : 16,
            }
        
        matplotlib.pyplot.figure('IV Curve',figsize=(6,4.5))
        ax = matplotlib.pyplot.subplot(1,1,1)
        ax.spines['bottom'].set_linewidth(1.3)
        ax.spines['left'].set_linewidth(1.3)
        ax.spines['top'].set_linewidth(1.3)        
        ax.spines['right'].set_linewidth(1.3)
        # 通过'k'指定线的颜色，lw指定线的宽度
        # 第三个参数除了颜色也可以指定线形，比如'r--'表示红色虚线
        # 更多属性可以参考官网：http://matplotlib.org/api/pyplot_api.html
        matplotlib.pyplot.plot(x, y, 'r', lw=1.5)
        matplotlib.pyplot.xlim(0, width)
        matplotlib.pyplot.xlabel('U (V)',fontdict=font)
        matplotlib.pyplot.ylim(0, height)
        matplotlib.pyplot.ylabel('I (mA/cm2)',fontdict=font)
        # scatter可以更容易地生成散点图
        #matplotlib.pyplot.scatter(x, y)
        annotate('area  = ' + str(format(area,".2f")+' cm2'),
                 xy=(0, 0), xycoords='data',
                 xytext=(+10, +110), textcoords='offset points', fontsize=12)    
        annotate('Isc  = ' + str(format(Isc,".2f")+' mA/cm2'),
                 xy=(0, 0), xycoords='data',
                 xytext=(+10, +85), textcoords='offset points', fontsize=12)
        annotate('Voc = ' + str(format(Voc,".2f")+'  V'),
                 xy=(0, 0), xycoords='data',
                 xytext=(+10, +60), textcoords='offset points', fontsize=12)
        annotate('FF   = ' + str(format(FF*100,".2f")+' %'),
                 xy=(0, 0), xycoords='data',
                 xytext=(+10, +35), textcoords='offset points', fontsize=12)
        annotate('Eff  = ' + str(format(eff,".2f")+'  %'),
                 xy=(0, 0), xycoords='data',
                 xytext=(+10, +10), textcoords='offset points', fontsize=12)
        matplotlib.pyplot.grid(False)
        # 将当前figure的图保存到文件
        # 将结果汇总
        if docupath:
            xlsfile_path = docupath+'/汇总.xls'
            if not os.path.exists(xlsfile_path):
                file = xlwt.Workbook()  #创建一个工作簿
                table = file.add_sheet('sheet 1')  #创建一个工作表
                table.write(0,0,'文件路径')
                table.write(0,1,'Jsc mA/cm2')
                table.write(0,2,'Voc V')
                table.write(0,3,'FF')
                table.write(0,4,'Eff %')
                table.write(0,5,'rsh')
                table.write(0,6,'rs')
                file.save(xlsfile_path)
                print('汇总文件'+xlsfile_path+'创建成功!') 
            else:
                print('汇总文件已存在')    
            
            bkall = xlrd.open_workbook(xlsfile_path) #获取表格中已有行数
            nxlsfile_path = xlsfile_path.replace(".xls",'')
            shall = bkall.sheet_by_index(0)
            irow = shall.nrows
            print("汇总表格中已有"+str(irow)+"行")
            bkcopyall=copy(bkall)
            shcopyall=bkcopyall.get_sheet(0)
            try:
                #style = xlwt.easyxf(num_format_str='#,##0.00')
                shcopyall.write(irow,0,nowdir)
                shcopyall.write(irow,1,Isc)
                shcopyall.write(irow,2,Voc)
                shcopyall.write(irow,3,FF)
                shcopyall.write(irow,4,eff)
                bkcopyall.save(xlsfile_path)  #保存
            except:
                print(nowdir+"汇总失败！")
        # 删除xls文档
        os.remove(nfile_path)
        print("xls文件删除成功！")
        
        nnfile_path = nowdir.replace(".csv",".png")
        print(nnfile_path+"图片保存成功！")
        matplotlib.pyplot.savefig(nnfile_path, bbox_inches='tight', dpi=150)
        return (nowdir,Isc,Voc,FF,eff)
    except:
        return
    
def documentselectinone():
    global area
    global width
    global height
    global Dark
    xall = []
    yall = []
    Iscall = []
    Vocall = []
    FFall = []
    Effall =[]
    name = []
    color = ["black","b","r","g","purple","olive","chocolate","deepskyblue","darkorange","lime","grey","royalblue"]
    cnames = {
        'aliceblue':            '#F0F8FF',
        'antiquewhite':         '#FAEBD7',
        'aqua':                 '#00FFFF',
        'aquamarine':           '#7FFFD4',
        'azure':                '#F0FFFF',
        'beige':                '#F5F5DC',
        'bisque':               '#FFE4C4',
        'black':                '#000000',
        'blanchedalmond':       '#FFEBCD',
        'blue':                 '#0000FF',
        'blueviolet':           '#8A2BE2',
        'brown':                '#A52A2A',
        'burlywood':            '#DEB887',
        'cadetblue':            '#5F9EA0',
        'chartreuse':           '#7FFF00',
        'chocolate':            '#D2691E',
        'coral':                '#FF7F50',
        'cornflowerblue':       '#6495ED',
        'cornsilk':             '#FFF8DC',
        'crimson':              '#DC143C',
        'cyan':                 '#00FFFF',
        'darkblue':             '#00008B',
        'darkcyan':             '#008B8B',
        'darkgoldenrod':        '#B8860B',
        'darkgray':             '#A9A9A9',
        'darkgreen':            '#006400',
        'darkkhaki':            '#BDB76B',
        'darkmagenta':          '#8B008B',
        'darkolivegreen':       '#556B2F',
        'darkorange':           '#FF8C00',
        'darkorchid':           '#9932CC',
        'darkred':              '#8B0000',
        'darksalmon':           '#E9967A',
        'darkseagreen':         '#8FBC8F',
        'darkslateblue':        '#483D8B',
        'darkslategray':        '#2F4F4F',
        'darkturquoise':        '#00CED1',
        'darkviolet':           '#9400D3',
        'deeppink':             '#FF1493',
        'deepskyblue':          '#00BFFF',
        'dimgray':              '#696969',
        'dodgerblue':           '#1E90FF',
        'firebrick':            '#B22222',
        'floralwhite':          '#FFFAF0',
        'forestgreen':          '#228B22',
        'fuchsia':              '#FF00FF',
        'gainsboro':            '#DCDCDC',
        'ghostwhite':           '#F8F8FF',
        'gold':                 '#FFD700',
        'goldenrod':            '#DAA520',
        'gray':                 '#808080',
        'green':                '#008000',
        'greenyellow':          '#ADFF2F',
        'honeydew':             '#F0FFF0',
        'hotpink':              '#FF69B4',
        'indianred':            '#CD5C5C',
        'indigo':               '#4B0082',
        'ivory':                '#FFFFF0',
        'khaki':                '#F0E68C',
        'lavender':             '#E6E6FA',
        'lavenderblush':        '#FFF0F5',
        'lawngreen':            '#7CFC00',
        'lemonchiffon':         '#FFFACD',
        'lightblue':            '#ADD8E6',
        'lightcoral':           '#F08080',
        'lightcyan':            '#E0FFFF',
        'lightgoldenrodyellow': '#FAFAD2',
        'lightgreen':           '#90EE90',
        'lightgray':            '#D3D3D3',
        'lightpink':            '#FFB6C1',
        'lightsalmon':          '#FFA07A',
        'lightseagreen':        '#20B2AA',
        'lightskyblue':         '#87CEFA',
        'lightslategray':       '#778899',
        'lightsteelblue':       '#B0C4DE',
        'lightyellow':          '#FFFFE0',
        'lime':                 '#00FF00',
        'limegreen':            '#32CD32',
        'linen':                '#FAF0E6',
        'magenta':              '#FF00FF',
        'maroon':               '#800000',
        'mediumaquamarine':     '#66CDAA',
        'mediumblue':           '#0000CD',
        'mediumorchid':         '#BA55D3',
        'mediumpurple':         '#9370DB',
        'mediumseagreen':       '#3CB371',
        'mediumslateblue':      '#7B68EE',
        'mediumspringgreen':    '#00FA9A',
        'mediumturquoise':      '#48D1CC',
        'mediumvioletred':      '#C71585',
        'midnightblue':         '#191970',
        'mintcream':            '#F5FFFA',
        'mistyrose':            '#FFE4E1',
        'moccasin':             '#FFE4B5',
        'navajowhite':          '#FFDEAD',
        'navy':                 '#000080',
        'oldlace':              '#FDF5E6',
        'olive':                '#808000',
        'olivedrab':            '#6B8E23',
        'orange':               '#FFA500',
        'orangered':            '#FF4500',
        'orchid':               '#DA70D6',
        'palegoldenrod':        '#EEE8AA',
        'palegreen':            '#98FB98',
        'paleturquoise':        '#AFEEEE',
        'palevioletred':        '#DB7093',
        'papayawhip':           '#FFEFD5',
        'peachpuff':            '#FFDAB9',
        'peru':                 '#CD853F',
        'pink':                 '#FFC0CB',
        'plum':                 '#DDA0DD',
        'powderblue':           '#B0E0E6',
        'purple':               '#800080',
        'red':                  '#FF0000',
        'rosybrown':            '#BC8F8F',
        'royalblue':            '#4169E1',
        'saddlebrown':          '#8B4513',
        'salmon':               '#FA8072',
        'sandybrown':           '#FAA460',
        'seagreen':             '#2E8B57',
        'seashell':             '#FFF5EE',
        'sienna':               '#A0522D',
        'silver':               '#C0C0C0',
        'skyblue':              '#87CEEB',
        'slateblue':            '#6A5ACD',
        'slategray':            '#708090',
        'snow':                 '#FFFAFA',
        'springgreen':          '#00FF7F',
        'steelblue':            '#4682B4',
        'tan':                  '#D2B48C',
        'teal':                 '#008080',
        'thistle':              '#D8BFD8',
        'tomato':               '#FF6347',
        'turquoise':            '#40E0D0',
        'violet':               '#EE82EE',
        'wheat':                '#F5DEB3',
        'white':                '#FFFFFF',
        'whitesmoke':           '#F5F5F5',
        'yellow':               '#FFFF00',
        'yellowgreen':          '#9ACD32'}
    morecolor = []
    for key in cnames:
        morecolor.append(cnames[key])
    print(morecolor)    
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:
            list1=os.path.join(parent,filename)
            nowdir=list1.replace('\\','/')
            npath = parent.replace('\\','/')
            npath1 = npath + '/AllInOne.png'
            judge = 1
            for v in Dark:
                if v in nowdir:
                    judge -=1
            if '.csv' in nowdir and judge ==1:
                labelname = filename.replace(".csv",'')
                name.append(labelname)
                print(name)
                c=open(nowdir,"r") #以r的方式打开csv文件
                read=csv.reader(c)
                nfile_path = nowdir.replace(".csv",".xls")
                print(nfile_path)
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
                file.save(nfile_path)  #保存
                print("xls创建成功！")
            
            
                if '.xls' in nfile_path and '汇总' not in nfile_path:
                    data=xlrd.open_workbook(nfile_path)
                    sheet1=data.sheet_by_index(0)
                    if sheet1.nrows==600:
                        a=0
                        for i1 in range(600):
                            if float(sheet1.cell(i1,0).value)>0 and a == 0:
                                Isc1 = float(sheet1.cell(i1,1).value) - (float(sheet1.cell(i1,1).value) - float(sheet1.cell(i1-1,1).value))/(float(sheet1.cell(i1,0).value) - float(sheet1.cell(i1-1,0).value)) * float(sheet1.cell(i1,0).value)
                                Isc = Isc1 / area*1000*(-1)
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
                        eff = pmax/area*1000
                        Pmax=Isc*Voc
                        FF = eff/Pmax
                        bkcopy=copy(data)
                        shcopy=bkcopy.get_sheet(0)
                        shcopy.write(1,6,Isc)
                        shcopy.write(1,7,Voc)
                        shcopy.write(1,8,FF)
                        shcopy.write(1,9,eff)
                        bkcopy.save(nfile_path)  #保存
                        print("Isc=",format(Isc,".2f"),"mA/cm2 Voc=",format(Voc,".2f"),"V FF=",format(FF*100,".2f"),"% Eff=",format(eff,".2f"),"%")
                        Iscall.append(Isc)
                        Vocall.append(Voc)
                        FFall.append(FF)
                        Effall.append(eff)
                data = xlrd.open_workbook(nfile_path)
                sheet1 = data.sheet_by_index(0)
                x = []
                y = []
                for i in range(600):
                    x.append(float(sheet1.cell(i,0).value))
                    y.append(-float(sheet1.cell(i,1).value)/area*1000)
                xall.append(x)
                yall.append(y)
                # 删除xls文档
                os.remove(nfile_path)
                print("xls文件删除成功！")
    # 通过rcParams设置全局横纵轴字体大小
    matplotlib.pyplot.rcParams['font.sans-serif']=['Arial']
    matplotlib.rcParams['xtick.direction'] = 'in' 
    matplotlib.rcParams['ytick.direction'] = 'in' 
    matplotlib.rcParams['xtick.labelsize'] = 12
    matplotlib.rcParams['ytick.labelsize'] = 12
    font = {'family' : 'Arial',
        'color'  : 'black',
        'weight' : 'normal',
        'size'   : 16,
        }
            
    matplotlib.pyplot.figure('IV Curve',figsize=(6,4.5))
    ax = matplotlib.pyplot.subplot(1,1,1)
    ax.spines['bottom'].set_linewidth(1.3)
    ax.spines['left'].set_linewidth(1.3)
    ax.spines['top'].set_linewidth(1.3)        
    ax.spines['right'].set_linewidth(1.3)
    # 通过'k'指定线的颜色，lw指定线的宽度
    # 第三个参数除了颜色也可以指定线形，比如'r--'表示红色虚线
    # 更多属性可以参考官网：http://matplotlib.org/api/pyplot_api.html
    if len(name) < 13:
        lenth = len(name)
        for i in range(lenth):
            matplotlib.pyplot.plot(xall[i], yall[i], color[i], lw=2,label = name[i])
    else:
        lenth = len(name)
        for i in range(lenth):
            matplotlib.pyplot.plot(xall[i], yall[i], morecolor[i], lw=2,label = name[i])
    matplotlib.pyplot.xlim(0, width)
    matplotlib.pyplot.xlabel('U (V)',fontdict=font)
    matplotlib.pyplot.ylim(0, height)
    matplotlib.pyplot.ylabel('I (mA/cm2)',fontdict=font)
    # scatter可以更容易地生成散点图
    #matplotlib.pyplot.scatter(x, y)
    matplotlib.pyplot.grid(False)
    matplotlib.pyplot.legend()
    # 将当前figure的图保存到文件

            
    nnfile_path = nowdir.replace(".csv",".png")
    print(nnfile_path+"图片保存成功！")
    matplotlib.pyplot.savefig(npath1, bbox_inches='tight', dpi=300)
    matplotlib.pyplot.show()
    print(Dark)
#主程序
area = 27.04
width = 0.7
height = 40
nfile_path = ""
Dark = ["dark","DARK"]
docupath = ""

root = Tk()
root.title('IVCurveDrawer v1.4')
Area = StringVar()
Width = StringVar()
Height = StringVar()
dark = StringVar()
path = StringVar()
Label(root,text = "请输入电池片面积，默认为27.04 cm2").grid(row = 2, column = 0,columnspan = 2)
Entry(root, textvariable = Area).grid(row = 2, column = 2)
Button(root, text = "输入面积", command = getarea).grid(row = 2, column = 3)
Label(root,text = "目标路径:").grid(row = 0,column = 0,columnspan = 2)
Entry(root, textvariable = path).grid(row = 0, column = 2,columnspan = 2,sticky=W)
Button(root, text = "画单张IV曲线图", command = targetselect).grid(row = 1, column = 1,sticky=W)
Button(root, text = "画多张IV曲线图", command = documentselect).grid(row = 1, column = 2,sticky=W)
Button(root, text = "多张IV曲线图画在一起", command = documentselectinone).grid(row = 1, column = 3,sticky=W)
Label(root,text = "请输入作图电压上限，默认为0.7 V").grid(row = 3, column = 0,columnspan = 2)
Entry(root, textvariable = Width).grid(row = 3, column = 2)
Button(root, text = "输入电压", command = getvolt).grid(row = 3, column = 3)
Label(root,text = "请输入作图电流上限，默认为40 mA/cm2").grid(row = 4, column = 0,columnspan = 2)
Entry(root, textvariable = Height).grid(row = 4, column = 2)
Button(root, text = "输入电流", command = getelec).grid(row = 4, column = 3)
Label(root,text = "请输入暗电流曲线包含的特征名称（空格分隔），默认为dark,DARK").grid(row = 5, column = 0,columnspan = 2)
Entry(root, textvariable = dark).grid(row = 5, column = 2)
Button(root, text = "输入名称", command = getdark).grid(row = 5, column = 3)
root.mainloop()






