import csv
#文件对话框头文件
import tkinter as tk
from tkinter import filedialog
import os
import xlwt

#写入预备
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet("提取结果")

rootc=tk.Tk()
rootc.withdraw()
filepath = filedialog.askdirectory()
row_num=0
col_num=0
lock=0
if filepath == "":
    print("\n取消选择")
else:
    for root, dirs, files in os.walk(filepath):
        #print(root) #当前目录路径
        #print(dirs) #当前路径下所有子目录
        #print(files) #当前路径下所有非目录子文件
        for file in files:
            fileroute = filepath + "/" + file
            with open(fileroute, 'r') as f:
                reader = csv.reader(f)
                writetext = list(reader)
                print(writetext[7][1])
                if(lock==0):
                    writetext[15][2]="大板条码"
                    seqq=0
                    for elemm in writetext[15]:
                        worksheet.write(row_num, col_num, writetext[15][seqq])
                        col_num = col_num + 1
                        seqq=seqq+1
                    row_num = row_num + 1
                    lock=1

                for elem in writetext:
                    if("Q8900" in elem):
                        print(elem)
                        seq=0
                        col_num=0
                        for i in elem:
                            if(seq==2):
                                worksheet.write(row_num, col_num, writetext[7][1])
                            else:
                                worksheet.write(row_num, col_num, elem[seq])
                            seq=seq+1
                            col_num=col_num+1
                        row_num=row_num+1
    workbook.save('数据提取.xls')