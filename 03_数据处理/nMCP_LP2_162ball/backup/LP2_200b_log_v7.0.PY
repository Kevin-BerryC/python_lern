import xlsxwriter
import time
import re
import os
import sys
#------------------------------------------------
#  优化文件处理（只接受txt文件）   ，依次将单个txt测试数据转化成excel文件                           
#------------------------------------------------
Title = ['HAND','TYPE','CASE','Freq','Volt','State','Temp','Yeild','CaseT','TT','TNO','TARGET','UNIT','UPER','LOWER','DUT1','DUT2','DUT3']
MULTI_CASE= ['VDD1', 'VDD2', 'VDDQ', 'VDCA', 'tQH', 'tDQSS']
MULTI_CASE= ['PNUM','VAS']
ENVIR=[
    14,# N0数据属性数量
    3, # N1: 样品数量
    1, # N2: 生成不带limit的图片
    1, # N3: 生成报告后删除PLOT文件夹
]

rootdir = os.path.dirname(os.path.realpath(sys.argv[0]))	#获取python工作目录
Tempe = 0

def get_info(x):    #获取变量信息
    try:
        print(x,end=' ')
        print(type(x),len(x))
    except:
        print(type(x))
    else:
        return

def GetData(file, DATABANK):     # excel关键信息筛选
    with open(file) as com:                             # 读出所有行
        content = com.readlines()
        com.close()
        #print(len(content))

    INF_CUT = 0
    for i in range(len(content)):           #筛选LOG中的有效数据
        pack = content[i]
        if pack.find('#IDDLOG:')>0 or pack.find('#DCLOG:')>0 or pack.find('#ODTLOG:')>0 or pack.find('#ACSHMOO:')>0:
            pack    = pack.split(':')
            temp1   = pack[1].split()
            del pack[1]                 
            pack    = pack+temp1            #提取完整数据
            for j in range(len(pack)):      #清洗数据
                pack[j] = pack[j].strip()
                if '*' in pack[j] or 'NO'in pack[j]:     #超出量程
                    pack[j] = '-999999'
            #pack[1] = int(pack[1])          #句柄格式化
            DATABANK.append(pack)           #存入数据库
            #print(pack)
            INF_CUT=INF_CUT+1

        #elif pack.find('#TRAINING:')>0 :
        #    pack    = pack.split(':')
        #    temp1   = pack[1].split()
        #    del pack[1]                 
        #    pack    = pack+temp1            #提取完整数据
        #    pack.insert(2,'1PS')
        #    for j in range(len(pack)):      #清洗数据
        #        pack[j] = pack[j].strip()
        #        if '*' in pack[j] or 'NO'in pack[j]:     #超出量程
        #            pack[j] = '-999999'
        #    #pack[1] = int(pack[1])          #句柄格式化
        #    DATABANK.append(pack)           #存入数据库
        #    #print(pack)
        #    INF_CUT=INF_CUT+1

        elif content[i].find('#TNO:') > 0:            #筛选LOG中有效测试信息
            pack    = content[i].split(':')
            #print(pack)

            del pack[0:3]           #清洗数据
            del pack[4:11:2]        #清洗数据

            for j in range(len(pack)):      #清洗数据
                pack[j] = pack[j].strip()

            for j in range(INF_CUT):                  #遍历数据库
                for k in range(len(pack)):              #添加信息
                    DATABANK[-1-j].insert(1, pack[-(k+1)])
                DATABANK[-1-j].insert(5, Tempe)
            INF_CUT = 0
        elif pack.find('#TEMPERATURE:')>=0:
            Tempe    = int(pack.split(':')[-1].strip())
    x=0
    #print (DATABANK)
    for i in range(len(DATABANK)):               #检查数据一致性
        #print (DATABANK[i])
        if len(DATABANK[i]) != ENVIR[0]+ENVIR[1]:
            x+=1
            print (DATABANK[i])
            print (len(DATABANK[i]), len(DATABANK))

def Getinfo(file, DATABANK, INFOBANK):     # excel关键信息筛选
    with open(file) as com:                 # 读出所有行
        content = com.readlines()
        com.close()
        #print(len(content))

    NUM_pack = []                           #创建数据集合
    for i in DATABANK:
        if i[14] not in NUM_pack:
            NUM_pack.append(i[14])

    for i in range(len(content)):                   #遍历LOG
        if content[i].find('#TNO:') > 0:            #筛选LOG中有效测试信息
            pack    = re.split('[,:]',content[i])
            Inf_Tno = int(pack[1])                  #句柄格式化
            if Inf_Tno not in NUM_pack:
                temp = ['#FUNC', Inf_Tno, pack[3],'NULL','NULL']
                for j in range(8):
                    temp.append('0')
                DATABANK.append(temp)

    for i in range(len(content)):                   #遍历LOG
        if content[i].find('#TNO:') > 0:            #筛选LOG中有效测试信息
            pack    = re.split('[,:]',content[i])
            Inf_Tno = int(pack[1])                  #句柄格式化

            if Inf_Tno not in NUM_pack:
                for j in range(3):              #清洗数据
                    del pack[0]
                for j in range(3,7):            #清洗数据
                    del pack[j]
                for j in range(len(pack)):      #清洗数据
                    pack[j] = pack[j].strip()

                for j in range(len(DATABANK)):                  #遍历数据库
                    if Inf_Tno == DATABANK[j][1]:               #找到句柄
                        for k in range(len(pack)):              #添加信息
                            DATABANK[j].insert(1, pack[-(k+1)])
                        #print(DATABANK[j])

    #x=0
    #for i in range(len(DATABANK)):               #检查数据一致性
    #    if len(DATABANK[i]) != 20:
    #        x+=1
    #        for j in range(20,len(DATABANK[i])):    #删除多余数据
    #            del DATABANK[i][16]
    #        print (DATABANK[i])
    #        print (len(DATABANK[i]))
    #    if DATABANK[i] == DATABANK[-1] and x>0:
    #        print('存在重复测试项%d,请检查Test flow' %x)

    for i in range(len(DATABANK)):              #信息二次细分
        if len(DATABANK[i]) < 26:
            pack = DATABANK[i][3].split('_')
            del DATABANK[i][3]
            if len(pack) == 7:          #确认格式一致后插入数据
                for j in range(len(pack)):
                    DATABANK[i].insert(3, pack[-(j+1)])
            #DATABANK[i].insert(16, 'NULL')

    Temp = []
    for i in range(len(content)):           #筛选LOG中的有效数据
        pack = content[i]

        if pack.find('DUT')==0:
            pack    = pack.split(':')
            if len(pack)==2:    continue
            temp   = pack[0].split()             #信息提取
            temp.append(pack[1])
            pack    = re.split('[,:]',pack[2])   #信息提取
            temp    += pack                      #信息提取
            for j in range(len(temp)):          #清洗数据
                temp[j] = temp[j].strip()
            Temp.append(temp)

        elif pack.find('#TNO:')>0 and len(Temp)>0:
            pack    = re.split('[,:]',content[i])
            INFOBANK.append(pack)
            INFOBANK += Temp
            Temp.clear()

    return

def Data_unify(DATABANK):           #数据单位统一
    x=0
    y=0
    z=0
    for i in range(len(DATABANK)):      #遍历数据库
        pack = DATABANK[i]
        Unit_L = re.findall('[A-Z]+', pack[ENVIR[0]-1])
        Unit_H = re.findall('[A-Z]+', pack[ENVIR[0]-2])
        if Unit_L ==['NULL'] and Unit_H ==['NULL']:         #不存在上下限赋为0
            for j in range(ENVIR[0]-2, ENVIR[0]+ENVIR[1]):
                DATABANK[i][j] = Trans(DATABANK[i][j], Unit_L)  #数据分离，单位换算
            DATABANK[i].insert(ENVIR[0]-2, 'NULL')

        elif Unit_L and Unit_L !=['NULL']:                  #优先处理下限单位
            x+=1
            for j in range(ENVIR[0]-2, ENVIR[0]+ENVIR[1]):
                DATABANK[i][j] = Trans(DATABANK[i][j], Unit_L)  #数据分离，单位换算
            DATABANK[i].insert(ENVIR[0]-2, Unit_L[0])                   #单位保留
        elif Unit_H and Unit_H !=['NULL']:                      #处理上限单位
            y+=1
            for j in range(ENVIR[0]-2, ENVIR[0]+ENVIR[1]):
                DATABANK[i][j] = Trans(DATABANK[i][j], Unit_H)   #数据分离，单位换算
            DATABANK[i].insert(ENVIR[0]-2, Unit_H[0])
        else:
            z+=1
    print(x,y,z)       #取单位下限个数，上限个数，没有传入limit的个数

def Trans(Raw, Unit):           #单位换算，返回数据
    unit = re.findall('[A-Z]+', Raw)
    data = re.sub('[A-Z]', '', Raw)

    if Unit == unit:
        if len(data)==0:
            return 0
        return float(data)
    elif data and float(data) == 0:
        return 0
    elif unit and unit[0] == 'NULL':
        return 0
    elif data and unit:
        if unit[0] == 'V' and Unit[0] == 'MV':
            return float(data)*1000
        if unit[0] == 'MV' and Unit[0] == 'V':
            return float(data)/1000
        if unit[0] == 'MV' and Unit[0] == 'NULL':
            return float(data)
        if unit[0] == 'V' and Unit[0] == 'NULL':
            return float(data)
        elif unit[0] == 'NA' and Unit[0] == 'UA':
            return float(data)/1000
        elif unit[0] == 'UA' and Unit[0] == 'MA':
            return float(data)/1000
        elif unit[0] == 'NA' and Unit[0] == 'MA':
            return float(data)/1000000
        elif unit[0] == 'S' and Unit[0] == 'MS':
            return float(data)*1000
        elif unit[0] == 'NS' and Unit[0] == 'PS':
            return float(data)*1000
        elif unit[0] == 'PS' and Unit[0] == 'NS':
            return float(data)/1000
        elif unit[0] == 'NS' and Unit[0] == 'MS':
            return float(data)/1000000
        elif unit[0] == 'NS' and Unit[0] == 'US':
            return float(data)/1000
        elif unit[0] == 'MS' and Unit[0] == 'US':
            return float(data)*1000
        elif unit[0] == 'KR' and Unit[0] == 'R':
            return float(data)*1000
        else :
            print(Raw, unit, Unit )      #检测未完成转换项目
    else :
        #print(1,unit,data)
        return float(Raw)

def Data_transnum(OP):
    for i in range(len(DATABANK)):
        try:    DATABANK[i][OP[0]]=int(DATABANK[i][OP[0]]) 
        except: pass
        try:    DATABANK[i][OP[1]]=int(DATABANK[i][OP[1]]) 
        except: pass
        try:    DATABANK[i][OP[2]]=int(DATABANK[i][OP[2]]) 
        except: pass
        if DATABANK[i][4]=='LV':    DATABANK[i][4]='1LV'
        if DATABANK[i][4]=='NV':    DATABANK[i][4]='2NV'
        if DATABANK[i][4]=='HV':    DATABANK[i][4]='3HV'

        if DATABANK[i][2] in MULTI_CASE: DATABANK[i][2]=DATABANK[i][ENVIR[0]-3]     #分离单个测试项多个参数的标题

def PutDATA(newfile, DATABANK, INFOBANK):
    newbook = xlsxwriter.Workbook(newfile)          # 生成xlsx
    wsheet = newbook.add_worksheet('DATA')          # 生成sheet
   #wsheet1 = newbook.add_worksheet('Device')          # 生成sheet

    #for i in range(len(DATABANK)):              #作图优化，清除频率单位，增加电压排序
    #    if 'RV' == DATABANK[i][9]:
    #        DATABANK[i][9] = DATABANK[i][9].replace('RV','2RV')
    #    elif 'LV' == DATABANK[i][9]:
    #        DATABANK[i][9] = DATABANK[i][9].replace('LV','1LV')
    #    elif 'HV' == DATABANK[i][9]:
    #        DATABANK[i][9] = DATABANK[i][9].replace('HV','3HV')

    wsheet.write_row(0,0, Title)
    for i in range(len(DATABANK)):
        wsheet.write_row(i+1, 0, DATABANK[i])    # 打印数据
    #for i in range(len(INFOBANK)):
    #    wsheet1.write_row(i, 0, INFOBANK[i])    # 打印数据
    newbook.close()  # 保存新生成的Excel

if __name__ == '__main__':
    start = time.perf_counter()
    filepath = 'D:\\DATA\\INPUT\\'
    name = os.listdir(filepath)

    for i in range(len(name)):
        if '.txt' in name[i]:
            print(name[i])
            DATABANK = []
            INFOBANK = []
            newfile = filepath+name[i][0:-4]+'.xlsx'
            RAW_file = filepath + name[i]
            GetData(RAW_file, DATABANK)
            #Getinfo(RAW_file, DATABANK, INFOBANK)
            Data_unify(DATABANK)
            Data_transnum([3,6,9])
            PutDATA(newfile, DATABANK, INFOBANK)
    end = time.perf_counter()
    print('程序执行时间: %.5s秒' % (end-start))
