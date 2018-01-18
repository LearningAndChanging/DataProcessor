import os
import os.path
import struct
from tkinter import *
from tkinter.filedialog import askdirectory 
import glob

slash = '/'
 
def change(a):
	file_or = open(rootdir + slash + a,'rb');
	file_de = open('D:/gif/' + a + '.gif','wb');
	first_200bytes = file_or.read(200);
 
	index = 0
	for byte_value in first_200bytes:
		if index == 1:
			if byte_value % 16 % 2 == 0:
				file_de.write(struct.pack('B',byte_value + 1))
				pass
			else:
				file_de.write(struct.pack('B',byte_value - 1))
				pass
		else:
			file_de.write(struct.pack('B',byte_value))
			pass
		index = (0 if index == 1 else 1)
		pass
	# file_de.write(b'GIF89a\xc8\x01');
	file_de.write(file_or.read());
	pass
 
 
'''
if(os.path.exists(modified_dir)):
	if(not os.path.isdir(modified_dir)):
		os.remove(modified_dir)
		os.mkdir(modified_dir)
	pass
else:
	os.mkdir(modified_dir)
	pass

for filenames in os.walk(rootdir): # 遍历当前目录下所有gif文件
	for file in filenames[2]:
		if len(file) == 32 :
			change(file);
			#print(file)
	break
''' 
 
#print('操作成功完成！')
#input();

def selectPath():
    rootdir = askdirectory()
    path.set(rootdir)
    for parent,dirnames,filenames in os.walk(rootdir):    #三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
        #print(filenames)
        #for filename in filenames:
        #       list=os.path.join(parent,filename)
        for file in filenames:
                #print(filenames)
                if len(file) == 32:
                        file_or = open(parent + slash + file,'rb');
                        file_de = open('D:/gif/' + file + '.gif','wb');
                        first_200bytes = file_or.read(200);
                 
                        index = 0
                        for byte_value in first_200bytes:
                                if index == 1:
                                        if byte_value % 16 % 2 == 0:
                                                file_de.write(struct.pack('B',byte_value + 1))
                                                pass
                                        else:
                                                file_de.write(struct.pack('B',byte_value - 1))
                                                pass
                                else:
                                        file_de.write(struct.pack('B',byte_value))
                                        pass
                                index = (0 if index == 1 else 1)
                                pass
                        # file_de.write(b'GIF89a\xc8\x01');
                        file_de.write(file_or.read());
                        pass

root = Tk()
path = StringVar()
Label(root,text = "在D盘新建gif文件夹单击路径选择后选择gif存储的文件夹").grid(row = 0, column = 0, columnspan = 3)
Label(root,text = "目标路径:").grid(row = 2,column = 0)
Entry(root, textvariable = path).grid(row = 2, column = 1)
Button(root, text = "路径选择", command = selectPath).grid(row = 2, column = 2)
root.mainloop()
