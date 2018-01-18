import os
import numpy as np
os.environ[ 'HOME' ] = '/tmp/'
os.environ[ 'MPLCONFIGDIR' ] = '/tmp/'
import matplotlib
matplotlib.use('Agg')
import csv
from pylab import annotate
import pickle
import hashlib
import pymysql
import time
import sys

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

def drawonepng(nowdir, filename, width = 0.7, height = 40):#nowdir为csv文件的路径
    global ROOTNOW
    global AREA
    global DARK
    
    Dark = ["dark", "Dark", "DARK", "-D.csv", "-d.csv", "-0.csv"]#暗电流曲线路径可能包含的特征字符串
    judge = 1
    for v in Dark:
        if v in nowdir:
            judge -=1 #判断暗电流曲线
    if '.csv' in nowdir and judge ==1:
        c = open(nowdir,"r") #以r的方式打开csv文件
        read = csv.reader(c)
        csvtmp = [line for line in read]
        volts = [float(csvtmp[i][2]) for i in range(47,647)]
        amps = [float(csvtmp[j][3]) for j in range(47,647)]
    try:
        a = 0
        for i1 in range(600):
            if volts[i1]>0 and a == 0:
                Isc1 = amps[i1] - (amps[i1] - amps[i1-1])/(volts[i1] - volts[i1-1]) * volts[i1]
                Isc = Isc1 / AREA*1000*(-1)
                a+=1
        b=0
        for i2 in range(600):
            if amps[i2]>0 and b == 0:
                Voc = volts[i2] - (volts[i2] - volts[i2-1])/(amps[i2] - amps[i2-1]) * amps[i2] 
                b+=1
        pmax=0
        for i3 in range(600):
            if volts[i3]*amps[i3]*(-1) > pmax:
                pmax = volts[i3]*amps[i3]*(-1)
        eff = pmax/AREA*1000
        Pmax = Isc*Voc
        FF = eff/Pmax
        #print("Isc=",format(Isc,".2f"),"mA/cm2 Voc=",format(Voc,".2f"),"V FF=",format(FF*100,".2f"),"% Eff=",format(eff,".2f"),"%")
        #if Isc<0 or Isc>40 or Voc<0 or Voc>0.7 or FF<0 or FF>0.9 or eff<5 or eff>20:
            #return
        x = volts
        y = [-amps[i]/AREA*1000 for i in range(600)]
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
        
        matplotlib.pyplot.figure('IV Curve', figsize=(6,4.5))
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
        matplotlib.pyplot.xlabel('U (V)', fontdict=font)
        matplotlib.pyplot.ylim(0, height)
        matplotlib.pyplot.ylabel('I (mA/cm2)', fontdict=font)
        # scatter可以更容易地生成散点图
        #matplotlib.pyplot.scatter(x, y)
        annotate('AREA  = ' + str(format(AREA,".2f")+' cm2'),
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
        #将当前figure的图保存到文件
        matplotlib.pyplot.savefig(ROOTNOW + "/pythonscript/png/IVCurve" + ".png", bbox_inches='tight', dpi=100)
        #print("图片保存成功！")
        return (nowdir, Isc, Voc, FF, eff)
    except:
        return

def calcuone(nowdir, filename):#nowdir为csv文件的路径
    global ROOTNOW
    global AREA
    global DARK
    
    Dark = ["dark", "Dark", "DARK", "-D.csv", "-d.csv", "-0.csv"]#暗电流曲线路径可能包含的特征字符串
    judge = 1
    for v in Dark:
        if v in nowdir:
            judge -=1 #判断暗电流曲线
    if '.csv' in nowdir and judge ==1:
        c = open(nowdir,"r") #以r的方式打开csv文件
        read = csv.reader(c)
        csvtmp = [line for line in read]
        volts = [float(csvtmp[i][2]) for i in range(47,647)]
        amps = [float(csvtmp[j][3]) for j in range(47,647)]
    try:
        a = 0
        for i1 in range(600):
            if volts[i1]>0 and a == 0:
                Isc1 = amps[i1] - (amps[i1] - amps[i1-1])/(volts[i1] - volts[i1-1]) * volts[i1]
                Isc = Isc1 / AREA*1000*(-1)
                a+=1
        b=0
        for i2 in range(600):
            if amps[i2]>0 and b == 0:
                Voc = volts[i2] - (volts[i2] - volts[i2-1])/(amps[i2] - amps[i2-1]) * amps[i2] 
                b+=1
        pmax=0
        for i3 in range(600):
            if volts[i3]*amps[i3]*(-1) > pmax:
                pmax = volts[i3]*amps[i3]*(-1)
        eff = pmax/AREA*1000
        Pmax = Isc*Voc
        FF = eff/Pmax
        return (nowdir, Isc, Voc, FF, eff)
    except:
        return
        
def drawall(width = 0.7, height = 40):
    global AREA
    global DARK
    global ROOTNOW
    
    xall = []
    yall = []
    Iscall = []
    Vocall = []
    FFall = []
    Effall =[]
    name = []
    color = ["black", "b", "r", "g", "purple", "olive", "chocolate", "deepskyblue", "darkorange", "lime", "grey", "royalblue"]
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
    #print(morecolor)    
    rootdir = ROOTNOW + "/upload/all"
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:
            nowdir = os.path.join(parent,filename).replace('\\','/')
            judge = 1
            for v in Dark:
                if v in nowdir:
                    judge -=1
            if '.csv' in nowdir and judge ==1:
                labelname = filename.replace(".csv",'')
                name.append(labelname)
                #print(name)
                c = open(nowdir,"r") #以r的方式打开csv文件
                read = csv.reader(c)
                csvtmp = [line for line in read]
                volts = [float(csvtmp[i][2]) for i in range(47,647)]
                amps = [float(csvtmp[j][3]) for j in range(47,647)]    
                x = volts
                y = [-amps[i]/AREA*1000 for i in range(600)]
                xall.append(x)
                yall.append(y)

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
            matplotlib.pyplot.plot(xall[i], yall[i], color[i], lw=2, label = name[i])
    else:
        lenth = len(name)
        for i in range(lenth):
            matplotlib.pyplot.plot(xall[i], yall[i], morecolor[i], lw=2, label = name[i])
    matplotlib.pyplot.xlim(0, width)
    matplotlib.pyplot.xlabel('U (V)', fontdict = font)
    matplotlib.pyplot.ylim(0, height)
    matplotlib.pyplot.ylabel('I (mA/cm2)', fontdict = font)
    matplotlib.pyplot.grid(False)
    matplotlib.pyplot.legend()
    # 将当前figure的图保存到文件
    matplotlib.pyplot.savefig(ROOTNOW + "/pythonscript/png/IVCurve" + ".png", bbox_inches='tight', dpi=100)

def drawone():
    global ROOTNOW
    global DARK
    
    rootdir = ROOTNOW + "/upload/all"
    csvmanager = CSVManager()
    
    for parent, dirnames, filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        
            nowdir = os.path.join(parent,filename).replace('\\','/')
            #print(nowdir)
            judge = 1
            for v in Dark:
                if v in nowdir:
                    judge -=1
                    #print("暗电流曲线")
            if '.csv' in nowdir and judge == 1:
                data = drawonepng(nowdir, filename)
                matplotlib.pyplot.close()
                csvmanager.add_new_csv(nowdir)
                if csvmanager.has_new_csv():
                    rt = csvmanager.get_new_csv()
                    if data:
                        try:
                            insert2db(data)
                        except:
                            print("数据已经存在于数据库中!")
                else:
                    print("数据已经存在于数据库中!")
    csvmanager.save_progress('new_csvs.txt',csvmanager.new_csvs)
    csvmanager.save_progress('old_csvs.txt',csvmanager.old_csvs)
    
def calcuall():
    global ROOTNOW
    global DARK
    
    rootdir = ROOTNOW + "/upload/all"
    csvmanager = CSVManager()
    
    for parent, dirnames, filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:                        
            nowdir = os.path.join(parent,filename).replace('\\','/')
            #print(nowdir)
            judge = 1
            for v in Dark:
                if v in nowdir:
                    judge -=1
                    #print("暗电流曲线")
            if '.csv' in nowdir and judge == 1:
                data = calcuone(nowdir, filename)
                matplotlib.pyplot.close()
                csvmanager.add_new_csv(nowdir)
                if csvmanager.has_new_csv():
                    rt = csvmanager.get_new_csv()
                    if data:
                        try:
                            insert2db(data)
                        except:
                            print("数据已经存在于数据库中!")
                else:
                    print("数据已经存在于数据库中!")
    csvmanager.save_progress('new_csvs.txt',csvmanager.new_csvs)
    csvmanager.save_progress('old_csvs.txt',csvmanager.old_csvs)
    
#主程序
ROOTNOW = sys.path[0].replace('\\','/')
Dark = ["dark", "Dark", "DARK", "-D.csv", "-d.csv", "-0.csv"]
f = open(ROOTNOW+"/tmp/parameter.txt")
AREA = float(f.readline())
f.close()

filenum = 0
for parent,dirnames,filenames in os.walk(ROOTNOW + "/upload/all"):
    for filename in filenames:
        filenum += 1
if filenum == 1:
    drawone()
elif filenum >= 1:
    drawall()
    #calcuall()
    

