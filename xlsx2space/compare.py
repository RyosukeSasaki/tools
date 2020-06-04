import os
import xlrd
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


def read_float_column(sheetname, label, column, index=None, flag=False):
    global datalist, indexlist
    data=wb.sheet_by_name(sheetname)
    mask = data.col_values(5)
    col = data.col_values(column)
    if index != None:
        ind = data.col_values(index)
    buf_data = []
    buf_index = []
    n=0
    for i in range(len(col)):
        if type(col[i]) is float:
            if (mask[i] == 0) or (flag == False):
                buf_data.append(col[i])
                if index != None:
                    buf_index.append(ind[i])
                else:
                    buf_index.append(n)
                n+=1
    datalist[label] = buf_data
    indexlist[label] = buf_index


try:
    read_float_column('Sheet1','na',6,0)
    read_float_column('Sheet1','no',7,0)
    read_float_column('Sheet1','wi',8,0)
    read_float_column('Sheet1','de',9,0)

except:
    pass



try:
    fig1=plt.figure(figsize=(15,10))
    g=fig1.add_subplot()
    g.yaxis.grid(linestyle='--', lw=1, alpha=0.6, color='black')
#    g.plot(indexlist['1_'], datalist["1_"], 'o', color='none', markersize=5, markeredgewidth=2, markeredgecolor='dodgerblue', alpha=0.8, label='A_n')
    g.plot(indexlist['na'], datalist['na'], alpha=0.8, label='パルス(狭)', color='darkgreen')
    g.plot(indexlist['no'], datalist['no'], alpha=0.8, label='パルス(中)', color='dodgerblue')
    g.plot(indexlist['wi'], datalist['wi'], alpha=0.8, label='パルス(広)', color='gold')
    g.plot(indexlist['de'], datalist['de'], alpha=0.8, label='δ関数', color='tomato')
    g.legend(loc='upper right', frameon=True, facecolor='none', edgecolor='lightgray',fontsize=14, markerscale=0.7)
    plt.savefig(outdir+'pulse_delta.png', bbox_inches="tight", pad_inches=0.05)
except:
    pass
finally:
    plt.clf()


#fig1=plt.figure(figsize=(15,10))
#g=fig1.add_subplot()
#g.yaxis.grid(linestyle='--', lw=1, alpha=0.6, color='black')
#g.bar(indexlist["p"],datalist["p"],width=0.6)
##g.legend(loc=4, frameon=True, facecolor='none', edgecolor='lightgray',fontsize=14, markerscale=0.7)
#plt.savefig(outdir+'phase.png', bbox_inches="tight", pad_inches=0.05)
#
#
#fig1=plt.figure(figsize=(15,10))
#g=fig1.add_subplot()
#g.yaxis.grid(linestyle='--', lw=1, alpha=0.6, color='black')
#g.bar(indexlist["A"],datalist["A"],width=0.6)
##g.legend(loc=4, frameon=True, facecolor='none', edgecolor='lightgray',fontsize=14, markerscale=0.7)
#plt.savefig(outdir+'Amp.png', bbox_inches="tight", pad_inches=0.05)
#plt.clf()
