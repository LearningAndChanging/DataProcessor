import xlrd
import xlwt
import os
import numpy as np
import matplotlib
from tkinter import *
from tkinter import filedialog
from pylab import annotate
from xlutils.copy import copy
from tkinter.filedialog import askdirectory

def getwavelength():
    global wavelength
    if Wavelength.get():
        s = Wavelength.get().split()
        wavelength+=s
    else:
        print("非法输入！")

def dict2list(dic:dict):
    ''' 将字典转化为列表 '''
    keys = dic.keys()
    vals = dic.values()
    lst = [(key, val) for key, val in zip(keys, vals)]
    return lst

def documentselectinone():
    global wavelength
    xall = []
    yall = []
    points_dict = {}
    nameall = []
    Absall = []
    name = []
    color = ["black","b","r","g","purple","olive","chocolate","deepskyblue","wavelengthorange","lime","grey","royalblue"]
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
        'wavelengthblue':             '#00008B',
        'wavelengthcyan':             '#008B8B',
        'wavelengthgoldenrod':        '#B8860B',
        'wavelengthgray':             '#A9A9A9',
        'wavelengthgreen':            '#006400',
        'wavelengthkhaki':            '#BDB76B',
        'wavelengthmagenta':          '#8B008B',
        'wavelengtholivegreen':       '#556B2F',
        'wavelengthorange':           '#FF8C00',
        'wavelengthorchid':           '#9932CC',
        'wavelengthred':              '#8B0000',
        'wavelengthsalmon':           '#E9967A',
        'wavelengthseagreen':         '#8FBC8F',
        'wavelengthslateblue':        '#483D8B',
        'wavelengthslategray':        '#2F4F4F',
        'wavelengthturquoise':        '#00CED1',
        'wavelengthviolet':           '#9400D3',
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
#    print(morecolor)    
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        for filename in filenames:
            list1=os.path.join(parent,filename)
            nowdir=list1.replace('\\','/')
            npath = rootdir.replace('\\','/')
            npath1 = npath + '/Lines.png'
            npath2 = npath + "/Points.png"
            if '.txt' in nowdir and "-2" not in nowdir:
                labelname = filename.replace(".txt",'')
                name.append(labelname)
                print(name)
                c=open(nowdir,"r") #以r的方式打开csv文件
                read=c.readlines()
                nfile_path = nowdir.replace(".txt",".xls")
                print("准备创建xls")
                file = xlwt.Workbook()  #创建一个工作簿
                table = file.add_sheet('sheet 1')  #创建一个工作表
                row = 0
                n = 0
                for line in read:
                    #print(len(line))
                    if row>1 and len(line)>10:
                        a=line[:6]
                        b=line[7:12]
                        #print(a)
                        #print(b)
                        table.write(n,0,a) #写入
                        table.write(n,1,b) 
                        n+=1
                    row+=1
                file.save(nfile_path)  #保存
                print(nfile_path+"创建成功！")
            
            
                if '.xls' in nfile_path and '.xlsx' not in nfile_path:
                    data=xlrd.open_workbook(nfile_path)
                    sheet1=data.sheet_by_index(0)
                    for i_row in range(sheet1.nrows):
                        if float(sheet1.cell(i_row,0).value) == float(wavelength[0]) and "H2O" not in nfile_path:
                            Abs = sheet1.cell(i_row,1).value
                            points_dict[float(filename.replace("min.txt",''))] = Abs
                    x = []
                    y = []
                    for i in range(sheet1.nrows):
                        x.append(float(sheet1.cell(i,0).value))
                        y.append(float(sheet1.cell(i,1).value))
                    xall.append(x)
                    yall.append(y)
                # 删除xls文档
                #os.remove(nfile_path)
                #print("xls文件删除成功！")
    sorted_dict = sorted(dict2list(points_dict), key = lambda asd:asd[0], reverse = False)
    for key,value in sorted_dict:
        nameall.append(key)
        Absall.append(value)
    #print(nameall)
    #print(Absall)
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
    matplotlib.pyplot.xlabel('Wavelenth (nm)',fontdict=font)
    matplotlib.pyplot.ylabel('Abs',fontdict=font)
    # scatter可以更容易地生成散点图
    #matplotlib.pyplot.scatter(x, y)
    matplotlib.pyplot.grid(False)
    matplotlib.pyplot.legend()
    # 将当前figure的图保存到文件        
    matplotlib.pyplot.savefig(npath1, bbox_inches='tight', dpi=300)
    #print(npath1+"图片保存成功！")
    matplotlib.pyplot.close()
    matplotlib.pyplot.plot(nameall, Absall, 'r', lw=1.5)
    matplotlib.pyplot.savefig(npath2, bbox_inches='tight', dpi=300)
    matplotlib.pyplot.close()
    #print(wavelength)


#主程序
wavelength = []
docupath = ""
nfile_path = ""

root = Tk()
root.title('CurveDrawer V1.0')
path = StringVar()
Wavelength = StringVar()

Label(root,text = "目标路径:").grid(row = 0,column = 0,columnspan = 2)
Entry(root, textvariable = path).grid(row = 0, column = 2,columnspan = 2,sticky=W)
Button(root, text = "作图", command = documentselectinone).grid(row = 0, column = 3,sticky=W)
Label(root,text = "请输入需要作图的波长").grid(row = 5, column = 0,columnspan = 2)
Entry(root, textvariable = Wavelength).grid(row = 5, column = 2)
Button(root, text = "输入名称", command = getwavelength).grid(row = 5, column = 3)
root.mainloop()






