import xlwt
import time
import re
import os


PINNUMB = []
PINLIST = []
SHEETNAME = ['SUM', 'OS_L', 'OS_H', 'CAODT', 'DQODT', 'PDDS', 'PUDS', 'IDD' ]
CaseNumber = [100, 101, 300, 307, 314, 321, 201]

DATABANK = []
for q in range(2000):
    DATABANK.append([0])

SUM_Array = [[]]
BANK_cunt = 0
FILE_cunt = 0

def Select_data(file):     # excel关键信息筛选
    data_array = []
    content = []
    with open(file, encoding='utf-8') as com:
        content = com.readlines()
        com.close()
    with open(file[:-20]+'.xls', 'a') as f:         #测试代码，不起作用，勿动
        for i in range(len(content)):
            aaa = ('#' + content[i]).strip()
            if aaa.find('DUT') > 0 or aaa.find('TEST NAME') > 0 or aaa.find('Test case') > 0:
                if aaa.find('UPPER') < 0 and aaa.find('DELAY') < 0:
                    for idx in range(0, len(PINNUMB)):
                        aaa = aaa.replace(PINNUMB[idx], PINLIST[idx])
                    data_array.append(aaa)
                    f.write(aaa+'\n')
    return data_array

def GetData(file, FILE_cunt, DATABANK, SUM_Array):     # excel关键信息筛选
    data_array = Select_data(file)          # 获取有效数据
    sumc = 0
    
    for q in range(2000):           # 遍历所有case容器
        for i in range(4):          # 创建4个DUT数据容器
            DATABANK[q].append({})  # 创建4个DUT数据容器
    for i in range(4):
        SUM_Array.append([])        # 创建4个SUM数据容器

    for i in range(len(data_array)):                            # 遍历数据
        data = data_array[i]                                    # 方便下一步处理
        if data.find('#TEST NAME') > 0:                               # SUM表内容提取
            data_name = re.findall(r'\d+', data)[0]
            dn = int(data_name)

        elif data.find('#Test case') > 0:
            continue
        elif data.find('PASS') > 0:
            continue
#            datatemp = re.findall(r'Result:(.*?), Y', data)
            Ntemp = re.findall(r'Title:(.*?), R', data)
            for dcunt in range(1+FILE_cunt*4,5+FILE_cunt*4):
                if datatemp[0][dcunt-FILE_cunt*4-1] == 'P':
                    aaa = 'PASS'
                elif datatemp[0][dcunt-FILE_cunt*4-1] == 'F':
                    aaa = 'FAIL'
                elif datatemp[0][dcunt-FILE_cunt*4-1] == '-':
                    aaa = 'NONE'
                SUM_Array[dcunt].append([])
                SUM_Array[dcunt][sumc].append(aaa)
            if FILE_cunt == 0:
                SUM_Array[0].append([])
                SUM_Array[0][sumc].append(Ntemp[0].strip())        # 获取测试名称
            sumc += 1
#
        elif data.find('FAIL') > 0:
            continue
        elif data.find('NO   TEST') > 0:
            continue
            
        else:
            dvalue = data.split()[1]                     # 获取测量值
            dup = data.split()[2]                        # 获取up limit
            dlow = data.split()[3]                       # 获取low limit
            dpin = data.split()[4]                       # 获取pin脚编号
            temp = data.split()[5]                       # 获取DUT
            temp = re.findall(r'\d+', temp)[0]
            ddut = int(temp)
            ddut = ddut + FILE_cunt * 4
            DATABANK[dn][ddut]['uper'] = dup
            DATABANK[dn][ddut]['lower'] = dlow
            DATABANK[dn][ddut][dpin] = dvalue

def PutDATA(newfile, DATABANK):
    wrrow = 1
    wrrow_cunt = 0
    num_temp = []
    complite = 0
    
    newbook = xlwt.Workbook(encoding="utf-8")
    wsheet = newbook.add_sheet(SHEETNAME[0])                # 生成sheet

    #wsheet.write(0, 0, 'Test Case')                         # 打印tilte             |    
    #wsheet.write(0, 1, 'Time')                              # 打印tilte             |    
    #for num in range(len(SUM_Array[0])):                    #                       |打印SUM表
    #    wsheet.write(wrrow, 0, SUM_Array[0][num][0])        #                       |
    #    wsheet.write(wrrow, 1, SUM_Array[0][num][1])        #                       |
    #    wrrow += 1                                          #                       |
    #for dut in range(1, len(SUM_Array)):                    #                       |
    #    wsheet.write(0, dut+1, 'DUT' + str(dut))            # 打印DUT编号            |
    #    wrrow = 1                                           #                       |
    #    for num in range(len(SUM_Array[dut])):              #                       |
    #        wsheet.write(wrrow, dut+1, SUM_Array[dut][num]) #                       |
    #        wrrow += 1                                      #                       |

    for num in range(len(DATABANK)):                        # --获取所有有效的case ---|    
        for dut in range(len(DATABANK[num])):               #                        |    
            if DATABANK[num][dut]:                          #                        |   
                if num not in num_temp:                     #                        |    
                    num_temp.append(num)                    #                        |   
    for num in num_temp:                                    # 依次处理case
        keylist = list(DATABANK[num][1].keys())             # 获取所有关键字           
        for i in range(len(CaseNumber)):                    #                          |判断是否需要生成SHEET
            if num == CaseNumber[i]:                        #                          |
                wsheet = newbook.add_sheet(SHEETNAME[i+1])  # 生成sheet                |
                wrrow_cunt = 0                              #                          |
                temp = 0                                    #                          |
                wsheet.write(0, 0, 'Test Case')             #打印title                  |
                wsheet.write(0, 1, 'Target')                #打印title                  |
                for dut in range(1, len(DATABANK[num])):    #打印title                  |
                    wsheet.write(0, dut+1, 'DUT'+str(dut))  #打印title                  |
        for dut in range(len(DATABANK[num])):               # 遍历所有DUT
            wrrow = 1                                       # 每次切换DUT都回到第一行开始打印
            for key in keylist:                             # 遍历所有字典
                #if num > 200 and key in ('uper','lower'):
                #    continue
                if DATABANK[num][dut]:                      # 不为空则继续处理
                    if num != complite:                     # 判断一个sheet内的case是否改变
                        complite = num                      # 改变先置位
                        wrrow_cunt += temp                  # 将本次打印占用的行数加上下一个测试项打印中
                    if dut == 1:
                        wsheet.write(wrrow+wrrow_cunt, dut-1, num)
                        wsheet.write(wrrow+wrrow_cunt, dut, key)
                    wsheet.write(wrrow+wrrow_cunt, dut+1, DATABANK[num][dut][key])
                    wrrow += 1
                    temp = len(keylist)                     #为打印下一个测试项做准备
                    #if num > 200:
                    #    temp = 1
    newbook.save(newfile)  # 保存新生成的Excel

if __name__ == '__main__':
    start = time.perf_counter()

    filepath = 'D:\\DATA\\'
    name = os.listdir(filepath)
    for i in name:
        #print(i)
        RAW_file = filepath + i
        newfile = filepath + i[:-10] +'.xls'
        #Select_data(RAW_file)
        GetData(RAW_file, FILE_cunt, DATABANK, SUM_Array)
        FILE_cunt += 1
    PutDATA(newfile, DATABANK)
    end = time.perf_counter()
    print('程序执行时间: %.5s秒' % (end-start))
