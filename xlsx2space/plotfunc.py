import xlrd
import os
import numpy as np
import math
import matplotlib.pyplot as plt
import argparse

plt.rcParams['font.family'] = 'Times New Roman'
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Noto Sans CJK JP']
plt.rcParams["font.size"]=18
path = os.getcwd()

parser = argparse.ArgumentParser()
parser.add_argument('name', help="ファイル名")
args=parser.parse_args()


datalist = {}
indexlist = {}

filedir=path+'/'+args.name

out = args.name.split('.')[0]
outdir=path+'/'+out+'/'
os.makedirs(outdir, exist_ok=True)

wb = xlrd.open_workbook(filedir)

x = np.arange(0,127,0.5)
y = list(map(lambda x: 324*math.cos(2*math.pi/127*5*x)+315*math.cos(2*math.pi/127*6*x-math.pi), x))

def read_float_column(sheetname, label, column, index=None):
    global datalist, indexlist
    data=wb.sheet_by_name(sheetname)
    col = data.col_values(column)
    if index != None:
        ind = data.col_values(index)
    buf_data = []
    buf_index = []
    n=0
    for i in range(len(col)):
        if type(col[i]) is float:
            buf_data.append(col[i])
            if index != None:
                buf_index.append(ind[i])
            else:
                buf_index.append(n)
            n+=1
    datalist[label] = buf_data
    indexlist[label] = buf_index


def read_float_row(sheetname, label, row, index=None):
    global datalist, indexlist
    data=wb.sheet_by_name(sheetname)
    col = data.row_values(row)
    if index != None:
        ind = data.row_values(index)
    buf_data = []
    buf_index = []
    n=0
    for i in range(len(col)):
        if type(col[i]) is float:
            buf_data.append(col[i])
            if index != None:
                buf_index.append(ind[i])
            else:
                buf_index.append(n)
            n+=1
    datalist[label] = buf_data
    indexlist[label] = buf_index


read_float_column('データ',"in",1,0)

try:
    fig1=plt.figure(figsize=(15,10))
    g=fig1.add_subplot()
    g.yaxis.grid(linestyle='--', lw=1, alpha=0.6, color='black')
    g.plot(indexlist["in"], datalist["in"], 'o', color='none', markersize=5, markeredgewidth=2, markeredgecolor='dodgerblue', alpha=0.8, label='入力')
    g.plot(x, y, alpha=0.8, label='理論理', color='darkgreen')
    g.legend(loc='lower right', frameon=True, facecolor='none', edgecolor='lightgray',fontsize=14, markerscale=0.7)
    plt.savefig(outdir+'theory.png', bbox_inches="tight", pad_inches=0.05)
except:
    pass
finally:
    plt.clf()

