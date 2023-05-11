import xlsxwriter
import time
import os
import docx
import sys

from docx.shared import Inches
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.shared import RGBColor 

import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt

Title = ['HAND','TYPE','CASE','BM','BL','Freq','ECC','RANK','Volt','State','TNO','TARGET']

rootdir = os.path.dirname(os.path.realpath(sys.argv[0]))+'\\'	#获取python工作目录

def get_info(x):    #获取变量信息
    try:
        print(x,'\n',type(x),len(x))
    except:
        print(x,'\n',type(x))
    else:
        return

def Make_rule(file,filepath):
    Rule = []
    rdata = pd.read_excel(file,sheet_name=0)
    #rdata = rdata.loc[rdata['HAND']!='#FUNC']   #屏蔽多个用法('#FUNC','#DCLOG')
    #get_info(rdata)

    for i in rdata.groupby('HAND'):
        Temp = [i[0],'','','','','','','']
        Rule.append(Temp)
        Temp = []
        for j in i[1].groupby('CASE'):                #利用CASE进行分组
            Temp.append(j[0])
        for j in range(int(len(Temp)/8)):
            Rule.append(Temp[0:8])
            del Temp[0:8]
        if len(Temp)>0:
            Rule.append(Temp)

    newbook = xlsxwriter.Workbook(filepath+'Rule.xlsx')         # 生成xlsx
    wsheet = newbook.add_worksheet('PLOT')                      # 生成sheet 
    wsheet1 = newbook.add_worksheet('WORD')                      # 生成sheet 

    style = newbook.add_format({
    'font': 'calibri',      # 字体
    'border': 1,            # 边框
    'bold': False,          # 字体加粗
    "align": "left",        # 对齐方式
    "valign": "vcenter",    # 垂直居中
    "color": "black",       # 字体颜色
    'fg_color': '#FFFFFF',  # 背景色
    })
    style1 = newbook.add_format({'fg_color': '#4F81BD','font': 'calibri','border': 1,'bold': 1,})

    wsheet.write_row(0, 0, Title, style1)
    for i in range(len(Rule)):
        if '#' in Rule[i][0]:
            wsheet.write_row(i+1, 0, Rule[i], style1)    # 打印数据
        else:
            wsheet.write_row(i+1, 0, Rule[i], style)    # 打印数据
    newbook.close()  # 保存新生成的Excel

def combine(file1, file2):
    F1 = pd.read_excel(file1,sheet_name=0)
    F1 = np.array(F1).tolist()
    F2 = pd.read_excel(file2,sheet_name=0)
    F2 = np.array(F1).tolist()

    for i in range(len(F1)):              #清除掉空值
        for j in range(len(F1[i])):
            if type(F1[i][j]) is float:
                del F1[i][j:]
                break
    for i in range(len(F2)):              #清除掉空值
        for j in range(len(F2[i])):
            if type(F2[i][j]) is float:
                del F2[i][j:]
                break
    
    for i in range(len(F1)):              #清除掉空值
        for j in range(len(F2)):
            if F2[j][14]==F1[i][14] and F2[j][15]==F1[i][15]:
                F1[i] = F1[i]+F2[j][-8:]
                print(F1[i])


def Ploat(file,filepath):
    rdata = pd.read_excel(file,sheet_name=0)
    rdata = pd.melt(rdata,id_vars=list(rdata)[0:-8])
    Rule = pd.read_excel(filepath+'Rule.xlsx',sheet_name='PLOT')
    Rule = np.array(Rule).tolist()

    plotpath = filepath[0:-6]+'PLOT\\'      #创建图片存放路径
    try:
        os.makedirs(plotpath)
        print('创建图片存放路径:',plotpath)
    except:
        print('图片存放路径已存在:',plotpath)

    for i in range(len(Rule)):              #清除掉空值
        for j in range(len(Rule[i])):
            if type(Rule[i][j]) is float:
                del Rule[i][j:]
                break

    matplotlib.use('Agg')
    #aa = matplotlib.get_backend()
    #get_info(aa)
    plt.figure(dpi=100,figsize=(24,14)) #制定画板大小
    for case in Rule:               #解析规则文件
        if len(case)==0:
            continue
        if '#' in case[0]:          #获取title 分类方式
            part = case[1:]
        elif 'dis' in case or '' in case :      #去掉注释行    
            #print(case)
            continue
        else:                       #获取内容开始画图
            for i in case:
                temp  = rdata.loc[rdata['CASE']==i]     #筛选当前处理case数据，如筛选多个用('#1','#2')
                pdata = temp.groupby(part)              #分类

                count = 0
                boxdata =[]
                label = []
                print(i,part)
                #print(temp)
                fig, ax1 = plt.subplots(figsize=(10, 4))

                for j in pdata:         #按照分类，添加数据到画布
                    count = count + 1
                    y = j[1]
                    y['value']=pd.to_numeric(y['value'])
                    y= y.dropna(axis=0,subset='value')  # delete NaN

                    AC_q1 = round(y['value'].quantile(q=0.25,interpolation="midpoint"),3)      #获取分位数
                    AC_q2 = round(y['value'].quantile(q=0.50,interpolation="midpoint"),3)
                    AC_q3 = round(y['value'].quantile(q=0.75,interpolation="midpoint"),3)
                    AC_robust_stdev = round((y['value'].quantile(q=0.75,interpolation="midpoint")-y['value'].quantile(q=0.25,interpolation="midpoint"))/1.35,3)
                    AC_high = round(AC_q2+6*AC_robust_stdev,3)
                    AC_low  = round(AC_q2-6*AC_robust_stdev,3)
                    plt.scatter(x=count,y=AC_high, c="b", marker='+')          #画出分布上下限
                    plt.scatter(x=count,y=AC_low, c="b", marker='+')

                    boxdata.append(y['value'])
                    label.append(j[0])

                    upper=(j[1].loc[j[1]['CASE']==i,'UPER'])
                    lower=(j[1].loc[j[1]['CASE']==i,'LOWER'])
                    unit =(j[1].loc[j[1]['CASE']==i,'UNIT'])
                    #get_info(unit)

                    try:                                                               #画上下限
                        if upper.iloc[0] != 0:
                            plt.scatter(x=count,y=upper.iloc[0], c="r", marker='_')
                        if lower.iloc[0] != 0:
                            plt.scatter(x=count,y=lower.iloc[0], c="r", marker='_')
                    except:
                        pass

                ax1.boxplot(boxdata, sym='o', patch_artist=1)   #作图
                bott = 0
                if len(label)>4:
                    ax1.set_xticklabels(label,rotation=90)      #添加标签以及设置显示角度
                    for l in label:
                        if len(l)>bott:
                            bott = len(l)
                            fig.subplots_adjust(left=0.10, right=0.97, top=0.92, bottom=0.18*bott)
                else:
                    ax1.set_xticklabels(label)                  #添加标签以及设置显示角度

                ax1.yaxis.grid(True)                            #添加水平线
                plt.xticks(fontsize=10)                         #标签字体大小
                plt.title(i,fontsize=16)                        #图片创建标题
                try:
                    plt.ylabel(unit.iloc[0],fontsize=10)    #拿到数值的单位
                except:
                    print('not found Unit',i)               #没有单位的情况下show出数据title

                i = i.replace("/", "_")                     #避免误识别，除号做下划线处理
                filename = plotpath+'boxplot_' + i + '.jpg' #生成文件名
                plt.savefig(filename)
                plt.clf()
                plt.close('all')
                #***********  第二张局域图  ***********
                '''fig, ax1 = plt.subplots(figsize=(10, 4))        #画布
                ax1.boxplot(boxdata, sym='o', patch_artist=1)   #作图
                if len(label)>4:            #自适应显示X轴标签旋转角度
                    ax1.set_xticklabels(label,rotation=90)      #添加标签以及设置显示角度
                    bott=0  
                    for l in label:         #自适应显示X轴标签预留空间
                        if len(l)>bott:
                            bott = len(l)
                            fig.subplots_adjust(left=0.10, right=0.97, top=0.92, bottom=0.18*bott)
                else:
                    ax1.set_xticklabels(label)                  #添加标签以及设置显示角度
                ax1.yaxis.grid(True)                            #添加水平线
                plt.xticks(fontsize=10)                         #标签字体大小
                plt.title(i,fontsize=16)                        #图片创建标题
                try:
                    plt.ylabel(unit.iloc[0],fontsize=10)    #拿到数值的单位
                except:
                    print('not found Unit',i)               #没有单位的情况下show出数据title
                filename = plotpath+'boxplot1_' + i + '.jpg' #生成文件名
                plt.savefig(filename)
                plt.clf()
                plt.close('all')'''

if __name__ == '__main__':
    start = time.perf_counter()

    name = os.listdir(rootdir)
    print(name)

    combine(rootdir+name[0],rootdir+name[1])
        #GetData(RAW_file, DATABANK)
        #Data_unify(DATABANK)
        #PutDATA(newfile, DATABANK)
        
    end = time.perf_counter()
    print('程序执行时间: %.5s秒' % (end-start))
