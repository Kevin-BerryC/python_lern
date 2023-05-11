import xlsxwriter
import time
import re 
import os

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

    x=0
    for i in range(len(DATABANK)):               #检查数据一致性
        if len(DATABANK[i]) != 20:
            x+=1
            for j in range(20,len(DATABANK[i])):    #删除多余数据
                del DATABANK[i][16]
            print (DATABANK[i])
            print (len(DATABANK[i]))
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

def Data_unify(DATABANK):           #数据单位统一
    x=0
    y=0
    z=0
    for i in range(len(DATABANK)):      #遍历数据库
        pack = DATABANK[i]
        Unit_L = re.findall('[A-Z]+', pack[17])
        Unit_H = re.findall('[A-Z]+', pack[16])

        if Unit_L ==['NULL'] and Unit_H ==['NULL']:         #不存在上下限赋为0
            DATABANK[i][17] = 0
            DATABANK[i][16] = 0
            DATABANK[i].insert(16, 'NULL')
        elif Unit_L and Unit_L !=['NULL']:                  #优先处理下限单位
            x+=1
            for j in range(16,26):
                DATABANK[i][j] = Trans(DATABANK[i][j], Unit_L)  #数据分离，单位换算
            DATABANK[i].insert(16, Unit_L[0])                   #单位保留
        elif Unit_H and Unit_H !=['NULL']:                      #处理上限单位
            y+=1
            for j in range(16,26):
                DATABANK[i][j] = Trans(DATABANK[i][j], Unit_H)   #数据分离，单位换算
            DATABANK[i].insert(16, Unit_H[0])
        else:
            z+=1
    print(x,y,z)       #取单位下限个数，上限个数，没有传入limit的个数

def Trans(Raw, Unit):           #单位换算，返回数据
    unit = re.findall('[A-Z]+', Raw)
    data = re.sub('[A-Z]', '', Raw)

    if Unit == unit:
        #print(data)
        return float(data)
    elif data and float(data) == 0:
        return 0
    elif unit and unit[0] == 'NULL':
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
            print(Raw)      #检测未完成转换项目
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
    filepath = 'D:\\DATA\\INPUT\\'
    name = os.listdir(filepath)
    
    for i in range(len(name)):
        #print(name[i])
        DATABANK = []
        newfile = filepath[0:-6]+name[i][0:-4]+'.xlsx'
        RAW_file = filepath + name[i]
        GetData(RAW_file, DATABANK)
        Data_unify(DATABANK)
        PutDATA(newfile, DATABANK)
    end = time.perf_counter()
    print('程序执行时间: %.5s秒' % (end-start))
