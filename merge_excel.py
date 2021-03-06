# -*- coding: utf-8 -*-

#将多个Excel文件合并成一个
import xlrd
import xlsxwriter
import os
from collections import defaultdict

    
# 获取目录中所有的excel文件
def get_excels():
    path = os.path.join(os.getcwd(), "excels")
    excels = [os.path.join(path, filename) for filename in os.listdir(path)]
    return excels

#打开一个excel文件
def open_xls(file):
    fh=xlrd.open_workbook(file)
    return fh

#获取excel中所有的sheet表
def getsheet(fh):
    return fh.sheets()

#获取sheet表的行数
def getnrows(fh,sheet):
    table=fh.sheets()[sheet]
    return table.nrows

#读取文件内容并返回行内容
def getFilect(file,shnum):
    fh=open_xls(file)
    table=fh.sheets()[shnum]
    num=table.nrows
    datavalue = []
    for row in range(num):
        rdata=table.row_values(row)
        datavalue.append(tuple(rdata))
    return datavalue

#获取sheet表的个数
def getshnum(fh):
    x=0
    sh=getsheet(fh)
    for sheet in sh:
        x+=1
    return x
    
    


if __name__=='__main__':
    #定义要合并的excel文件列表
    # allxls=['F:/test/excel1.xlsx','F:/test/excel2.xlsx']
    #存储所有读取的结果
    allxls = get_excels()
    datavalue = defaultdict(list)
    sheet_names_lst = []
    
    # 用字典存储每个对应sheet的内容，同时获取sheet名字
    for fl in allxls:
        fh=open_xls(fl)
        x=getshnum(fh)
        file_sheet_names = fh.sheet_names()
        sheet_names_lst.append(file_sheet_names)
        for shnum in range(x):
            print("正在读取文件："+str(fl)+"的第"+str(shnum)+"个sheet表的内容...")
            sheet_value = getFilect(fl,shnum)
            datavalue[shnum] += sheet_value
            # rvalue=getFilect(fl,shnum)
    #定义最终合并后生成的新文件
    sheet_names = sorted(sheet_names_lst, key=lambda x: len(x), reverse=True)[0] # 得到最终的sheet名字
    
    print(sheet_names)
    result_path = os.path.join(os.getcwd(), "result")
    merge_file = "merge_file.xlsx"
    # endfile='F:/test/excel3.xlsx'
    
    #创建一个sheet工作对象
    wb=xlsxwriter.Workbook(merge_file)
    ws_handles = [wb.add_worksheet(sheet_names[shnum]) for shnum in range(x)]
    
    
    for shnum in range(x):
        sheet_data = list(set(datavalue[shnum]))
        for a in range(len(sheet_data)):
            for b in range(len(sheet_data[a])):
                c=sheet_data[a][b]
                ws_handles[shnum].write(a,b,c)
                
    wb.close()
            

    print("文件合并完成")


