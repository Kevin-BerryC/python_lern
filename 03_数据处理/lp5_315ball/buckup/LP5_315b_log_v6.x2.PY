import xlsxwriter
import time
import re
import os
#import codecs
import os.path
#import zipfile
#import docx

#from docx.shared import Inches
#from docx.enum.section import WD_ORIENT
#from docx.enum.table import WD_TABLE_ALIGNMENT
#from docx.oxml import parse_xml
#from docx.oxml.ns import nsdecls
#from docx.oxml.ns import qn
#from docx.shared import Pt
#from docx.shared import RGBColor 
#from itertools import cycle
#from string import digits

import numpy as np
import pandas as pd
#import warnings
#warnings.filterwarnings("ignore")
import matplotlib.pyplot as plt
#import statsmodels.api as sm

# rootdir = os.path.dirname(os.path.realpath(sys.argv[0]))	#获取python工作目录

def get_info(x):    #获取变量信息
    try:
        print(x)
        print(type(x),len(x))
    except:
        print(type(x))
    else:
        return

def Ploat_rule(file,filepath):
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
    wsheet = newbook.add_worksheet('Rule')                      # 生成sheet

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
    for i in range(len(Rule)):
        if '#' in Rule[i][0]:
            wsheet.write_row(i+1, 0, Rule[i],style1)    # 打印数据
        else:
            wsheet.write_row(i+1, 0, Rule[i],style)    # 打印数据
    newbook.close()  # 保存新生成的Excel

def Ploat(file,filepath):
    rdata = pd.read_excel(file,sheet_name=0)
    Rule = pd.read_excel(filepath+'Rule.xlsx',sheet_name=0)
    Rule = np.array(Rule).tolist()

    for hand in rdata.groupby(['HAND']):
        data = pd.melt(rdata,id_vars=list(hand[1])[0:-8])
        CASE = hand[1].groupby('CASE')              #利用CASE进行分组
        name = (CASE.size().index).to_list()        #获取分组名称并转换为列表

        if 'DCLOG' in hand[0]:
            pdata = data.groupby(['CASE','Volt'])
            continue
        elif 'ACS' in hand[0]:
            pdata = data.groupby(['CASE','TARGET'])#(['CASE','TARGET','Freq'])
        else:   
            print('未处理项：',hand[0])
            continue
            #get_info(hand[0])

        plt.figure(dpi=100,figsize=(24,14))

        Plotcunt= 0
        for i in name:
            #continue
            count = 0
            data =[]
            label = []
            print(i)
            Plotcunt+=1
            fig, ax1 = plt.subplots(figsize=(10, 4))
            fig.canvas.manager.set_window_title('A Boxplot Example')
            fig.subplots_adjust(left=0.10, right=0.97, top=0.92, bottom=0.35)

            for j in pdata:
                x = list(j[0])      #名称转换
                if len(x) >1:       #有多个名称的取第一个，前面分类CASE必须排在第一位
                    x = x[0]        #取第一个
                if i == x:
                    #get_info(j)
                    #get_info(i)
                    count = count + 1
                    y = j[1]
                    y['value']=pd.to_numeric(y['value'])
                    y= y.dropna(axis=0,subset='value')  # delete NaN

                    #print(y['value'])
                    AC_q1 = round(y['value'].quantile(q=0.25,interpolation="midpoint"),3)
                    AC_q2 = round(y['value'].quantile(q=0.50,interpolation="midpoint"),3)
                    AC_q3 = round(y['value'].quantile(q=0.75,interpolation="midpoint"),3)
                    AC_robust_stdev = round((y['value'].quantile(q=0.75,interpolation="midpoint")-y['value'].quantile(q=0.25,interpolation="midpoint"))/1.35,3)

                    AC_high = round(AC_q2+6*AC_robust_stdev,3)
                    AC_low  = round(AC_q2-6*AC_robust_stdev,3)

                    plt.scatter(x=count,y=AC_high, c="b", marker='+')
                    plt.scatter(x=count,y=AC_low, c="b", marker='+')

                    data.append(y['value'])
                    label.append(j[0][1:])

                    upper=(j[1].loc[j[1]['CASE']==i,'UPER'])
                    lower=(j[1].loc[j[1]['CASE']==i,'LOWER'])
                    unit =(j[1].loc[j[1]['CASE']==i,'UNIT'])
                    #get_info(unit)
                    unit= unit.iloc[0]
                    #get_info(unit)

                    #try:
                    #    if upper.iloc[0] != 0:
                    #        plt.scatter(x=count,y=upper.iloc[0], c="r", marker='_')
                    #    if lower.iloc[0] != 0:
                    #        plt.scatter(x=count,y=lower.iloc[0], c="r", marker='_')
                    #except:
                    #    pass
            #continue
            ax1.boxplot(data)
            ax1.set_xticklabels(label,rotation=90, fontsize=6)

            i = i.replace("/", "_")
            filename='D:\\DATA\\INPUT\\'+'boxplot_' + i + '_AC.jpg'
            plt.xticks(fontsize=10)
            plt.title(i,fontsize=15)

            try:
                plt.ylabel(unit,fontsize=10) #拿到单位
            except:
                #get_info(unit)
                pass

            plt.savefig(filename)
            plt.clf()
            plt.close()
            #os._exit(0)

if __name__ == '__main__':
    start = time.perf_counter()
    filepath = 'D:\\DATA\\INPUT\\'
    name = os.listdir(filepath)

    for i in range(len(name)):
        #print(name[i])
        DATABANK = []
        newfile = filepath[0:-6]+name[i][0:-4]+'.xlsx'
        RAW_file = filepath + name[i]
        if '.xlsx' not in RAW_file:
            os.remove(RAW_file)
            continue
        elif 'Rule' in RAW_file:
            continue
        
        if 'Rule.xlsx' not in name:
            Ploat_rule(RAW_file,filepath)
            print('\n已生成规则文件,请根据需求前往调整:路径',filepath,'\n')
            break
        
        #GetData(RAW_file, DATABANK)
        #Data_unify(DATABANK)
        #PutDATA(newfile, DATABANK)
        Ploat(RAW_file,filepath)
        #os._exit(0)
    end = time.perf_counter()
    print('程序执行时间: %.5s秒' % (end-start))
