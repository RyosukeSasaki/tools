import xlrd
import os
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


ZEN = "".join(chr(0xff01 + i) for i in range(94))
HAN = "".join(chr(0x21 + i) for i in range(94))
HAN2ZEN = str.maketrans(HAN, ZEN)


datalist = {}
indexlist = {}

filedir=path+'/'+args.name

out = args.name.split('.')[0]
outdir=path+'/'+out+'/'
os.makedirs(outdir, exist_ok=True)

wb = xlrd.open_workbook(filedir)


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


read_float_column(str(0).translate(HAN2ZEN)+'次',"wave"+str(0),2)
for i in range(1,65):
    try:
        read_float_column(str(i).translate(HAN2ZEN)+'次',"wave"+str(i),i+3)
    except:
        continue

read_float_column('データ',"in",1,0)
read_float_row('データ',"p",6,0)
read_float_row('データ',"A",5,0)


for i in range(0,65):
    try:
        fig1=plt.figure(figsize=(15,10))
        g=fig1.add_subplot()
        g.yaxis.grid(linestyle='--', lw=1, alpha=0.6, color='black')
        g.plot(indexlist["in"], datalist["in"], 'o', color='none', markersize=5, markeredgewidth=2, markeredgecolor='dodgerblue', alpha=0.8, label='入力')
        g.plot(indexlist["wave"+str(i)], datalist["wave"+str(i)], alpha=0.8, label=str(i)+'次', color='darkgreen')
        g.legend(loc='lower right', frameon=True, facecolor='none', edgecolor='lightgray',fontsize=14, markerscale=0.7)
        plt.savefig(outdir+'wave'+str(i)+'.png', bbox_inches="tight", pad_inches=0.05)
    except:
        continue
    finally:
        plt.clf()


fig1=plt.figure(figsize=(15,10))
g=fig1.add_subplot()
g.yaxis.grid(linestyle='--', lw=1, alpha=0.6, color='black')
g.bar(indexlist["p"],datalist["p"],width=0.6)
#g.legend(loc=4, frameon=True, facecolor='none', edgecolor='lightgray',fontsize=14, markerscale=0.7)
plt.savefig(outdir+'phase.png', bbox_inches="tight", pad_inches=0.05)


fig1=plt.figure(figsize=(15,10))
g=fig1.add_subplot()
g.yaxis.grid(linestyle='--', lw=1, alpha=0.6, color='black')
g.bar(indexlist["A"],datalist["A"],width=0.6)
#g.legend(loc=4, frameon=True, facecolor='none', edgecolor='lightgray',fontsize=14, markerscale=0.7)
plt.savefig(outdir+'Amp.png', bbox_inches="tight", pad_inches=0.05)
plt.clf()
