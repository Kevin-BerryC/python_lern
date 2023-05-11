import xlsxwriter
import time
import re
import os
import pandas as pd

Title = ['HAND','TYPE','CASE','BM','BL','Freq','Ratio','ECC','RANK','Volt','State','Yeild','CaseT','TT','TNO','TARGET','UNIT','UPER','LOWER','DUT1','DUT2','DUT3','DUT4','DUT5','DUT6','DUT7','DUT8']

def GetData(file, DATABANK):     # excel关键信息筛选

    with open(file) as com:                             # 读出所有行
        content = com.readlines()
        com.close()
        #print(len(content))      

    for i in range(len(content)):           #筛选LOG中的有效数据
        pack = content[i]
        if pack.find('#IDDLOG:')>0 or pack.find('#DCLOG:')>0 or pack.find('#ODTLOG:')>0 or pack.find('#CASHMOO:')>0 or pack.find('#ACSHMOO:')>0:
            pack    = pack.split(':')
            temp1   = pack[1].split()
            del pack[1]
            pack    = pack+temp1            #提取完整数据
            for j in range(len(pack)):      #清洗数据
                pack[j] = pack[j].strip()
                if pack[j].find('*')>=0 or pack[j].find('NOPASS')>=0:     #超出量程
                    pack[j] = '-999999'
            pack[1] = int(pack[1])          #句柄格式化
            DATABANK.append(pack)           #存入数据库

    for i in range(len(content)):                   #遍历LOG
        if content[i].find('#TNO:') > 0:            #筛选LOG中有效测试信息
            pack    = re.split('[,:]',content[i])
            Inf_Tno = int(pack[1])                  #句柄格式化

            del pack[0:3]                   #清洗数据
            for j in range(3,7):            #清洗数据
                del pack[j]
            for j in range(len(pack)):      #清洗数据
                pack[j] = pack[j].strip()

            for j in range(len(DATABANK)):                  #遍历数据库
                if Inf_Tno == DATABANK[j][1]:               #找到句柄
                    for k in range(len(pack)):              #添加信息
                        DATABANK[j].insert(1, pack[-(k+1)])

    x=0
    for i in range(len(DATABANK)):               #检查数据一致性
        if len(DATABANK[i]) != 16:
            x+=1
            #for j in range(16,len(DATABANK[i])):    #删除多余数据
            #    print(DATABANK[i])
            #    del DATABANK[i][16]
            print ('长度不一致数据:',DATABANK[i],'长度:',len(DATABANK[i]))
        if DATABANK[i] == DATABANK[-1] and x>0:
            print('存在重复测试项%d,请检查Test flow' %x)

    for i in range(len(DATABANK)):              #信息二次细分
        pack = DATABANK[i][3].split('_')
        del DATABANK[i][3]

        if len(pack)== 2:           #DC测试项填充格式
            del pack[0]
            for j in range(6):
                pack.insert(0, 'NA')
        if len(pack) == 7:          #确认格式一致后插入数据
            for j in range(len(pack)):
                DATABANK[i].insert(3, pack[-(j+1)])
    return

def GetData_2(file, DATABANK, Filecut):     # excel关键信息筛选
    STAT_Tno = []
    for i in range(len(DATABANK)):
        if DATABANK[i][14] not in STAT_Tno:
            STAT_Tno.append(DATABANK[i][14])

    with open(file) as com:                             # 读出所有行
        content = com.readlines()
        com.close()
        #print(len(content))
    
    for i in range(len(content)):           #筛选LOG中的有效数据
        pack = content[i]
        if pack.find('#IDDLOG:')>0 or pack.find('#DCLOG:')>0 or pack.find('#ODTLOG:')>0 or pack.find('#CASHMOO:')>0 or pack.find('#ACSHMOO:')>0:
            pack    = pack.split(':')
            temp1   = pack[1].split()
            del pack[1]                 
            pack    = pack+temp1            #提取完整数据
            for j in range(len(pack)):                                  #清洗数据
                pack[j] = pack[j].strip()
                if pack[j].find('*')>=0 or pack[j].find('NOPASS')>=0:   #超出量程
                    pack[j] = '-999999'
            pack[1] = int(pack[1])          #句柄格式化

            if pack[1] not in STAT_Tno:         #检测差异项
                print('差异测试项',pack)
            for j in range(len(DATABANK)):      #对比相同测试项
                if pack[1]==DATABANK[j][14] and pack[2]==DATABANK[j][15]:
                    #print(pack,DATABANK[j])
                    del pack[0:5]
                    DATABANK[j]+=pack           #存入数据库

    for i in range(len(content)):                   #遍历LOG
        if content[i].find('#TNO:') > 0:            #筛选LOG中有效测试信息
            pack    = re.split('[,:]',content[i])
            for j in range(len(pack)):              #清洗数据
                pack[j] = pack[j].strip()
            pack[1] = int(pack[1])                  #句柄格式化

            for j in range(len(DATABANK)):                  #遍历数据库
                if pack[1]==DATABANK[j][14]:                #找到数据编号
                    DATABANK[j][10] = DATABANK[j][10] + '_' + pack[7]           #测试状态合并
                    DATABANK[j][11] = (float(DATABANK[j][11])*Filecut + float(pack[9]))/(100*Filecut) #测试良率合并
    #x=0
    #for i in range(len(DATABANK)):               #检查数据一致性
    #    if len(DATABANK[i]) != 16:
    #        x+=1
    #        for j in range(16,len(DATABANK[i])):    #删除多余数据
    #            print(DATABANK[i])
    #            del DATABANK[i][16]
    #        print (len(DATABANK[i]))
    #    if DATABANK[i] == DATABANK[-1] and x>0:
    #        print('存在重复测试项%d,请检查Test flow' %x)
#
    #for i in range(len(DATABANK)):              #信息二次细分
    #    pack = DATABANK[i][3].split('_')
    #    del DATABANK[i][3]
#
    #    if len(pack)== 2:           #DC测试项填充格式
    #        del pack[0]
    #        for j in range(6):
    #            pack.insert(0, 'NA')
    #    if len(pack) == 7:          #确认格式一致后插入数据
    #        for j in range(len(pack)):
    #            DATABANK[i].insert(3, pack[-(j+1)])
    return

def Data_unify(DATABANK):           #数据单位统一
    x=0
    y=0
    z=0
    for i in range(len(DATABANK)):      #遍历数据库
        pack = DATABANK[i]
        Unit_L = re.findall('[A-Z]+', pack[17])
        Unit_H = re.findall('[A-Z]+', pack[16])

        if Unit_H[0] == 'A':       #过大单位矫正
            #print(Unit_H)
            Unit_H[0] = 'MA'

        if Unit_L ==['NULL'] and Unit_H ==['NULL']:         #不存在上下限赋为0
            for j in range(16,len(DATABANK[i])):
                DATABANK[i][j] = Trans(DATABANK[i][j], Unit_L)  #数据分离，单位换算
            DATABANK[i].insert(16, Unit_L[0])
        elif Unit_L and Unit_L !=['NULL']:                  #优先处理下限单位
            x+=1
            for j in range(16,len(DATABANK[i])):
                DATABANK[i][j] = Trans(DATABANK[i][j], Unit_L)  #数据分离，单位换算
            DATABANK[i].insert(16, Unit_L[0])                   #单位保留
        elif Unit_H and Unit_H !=['NULL']:                      #处理上限单位
            y+=1
            for j in range(16,len(DATABANK[i])):
                DATABANK[i][j] = Trans(DATABANK[i][j], Unit_H)   #数据分离，单位换算
            DATABANK[i].insert(16, Unit_H[0])
        else:
            z+=1
    print('单位取下限个数%d'%x,'单位取上限个数%d'%y,'异常个数:%d'%z)       #取单位下限个数，上限个数，没有传入limit的个数

def Trans(Raw, Unit):           #单位换算，返回数据
    unit = re.findall('[A-Z]+', Raw)
    data = re.sub('[A-Z]', '', Raw)

    if unit and unit[0] == 'NULL':
        return 0
    elif Unit == unit:
        return float(data)
    elif data and float(data) == 0:
        return 0
    elif data and unit:
        if unit[0] == 'V' and Unit[0] == 'MV':
            return float(data)*1000
        elif unit[0] == 'NA' and Unit[0] == 'UA':
            return float(data)/1000
        elif unit[0] == 'UA' and Unit[0] == 'MA':
            return float(data)/1000
        elif unit[0] == 'NA' and Unit[0] == 'MA':
            return float(data)/1000000
        elif unit[0] == 'A' and Unit[0] == 'MA':
            return float(data)*1000
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
        elif unit[0] == 'KR' and Unit[0] == 'R':
            return float(data)*1000
        else :
            print(unit,Unit,data,Raw)
            #print(Raw)      #检测未完成转换项目
            return float(Raw)
    else :
        #print(1,unit,data)
        return float(Raw)

def PutDATA(newfile, DATABANK):

    newbook = xlsxwriter.Workbook(newfile)          # 生成xlsx
    wsheet = newbook.add_worksheet('DATA')          # 生成sheet

    wsheet.write_row(0,0, Title)
    for i in range(len(DATABANK)):
        for j in range(len(DATABANK[i])):
            wsheet.write(i+1, j, DATABANK[i][j])    # 打印数据
    newbook.close()  # 保存新生成的Excel

if __name__ == '__main__':
    start = time.perf_counter()
    DATABANK = []
    filepath = 'D:\\DATA\\INPUT\\'
    newfile = filepath[0:-6] + 'DATA.xlsx'
    name = os.listdir(filepath)

    for i in range(len(name)):
        print(name[i])
        RAW_file = filepath + name[i]
        if i == 0:
            GetData(RAW_file, DATABANK)
        else :
            GetData_2(RAW_file, DATABANK,i+1)
            pass
    Data_unify(DATABANK)
    PutDATA(newfile, DATABANK)
    end = time.perf_counter()
    print('程序执行时间: %.5s秒' % (end-start))
