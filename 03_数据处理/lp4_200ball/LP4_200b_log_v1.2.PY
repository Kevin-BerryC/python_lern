import xlwt
import time
import re
import os


PINNUMB = ['PIN1863',	'PIN1864',	'PIN1861',	'PIN1862',	'PIN97'  ,	'PIN289' ,	'PIN104' ,	'PIN292' ,	'PIN71'  ,	'PIN263' ,	'PIN79'  ,	'PIN271' ,	'PIN65'  ,	'PIN257' ,	'PIN72'  ,	'PIN264' ,	'PIN73'  ,	'PIN265' ,	'PIN80'  ,	'PIN272' ,	'PIN81'  ,	'PIN273' ,	'PIN82'  ,	'PIN274' ,	'PIN111' ,	'PIN335' ,	'PIN66'  ,	'PIN67'  ,	'PIN69'  ,	'PIN70'  ,	'PIN74'  ,	'PIN75'  ,	'PIN76'  ,	'PIN78'  ,	'PIN98'  ,	'PIN99'  ,	'PIN101' ,	'PIN102' ,	'PIN106' ,	'PIN107' ,	'PIN108' ,	'PIN110' ,	'PIN258' ,	'PIN259' ,	'PIN261' ,	'PIN262' ,	'PIN266' ,	'PIN267' ,	'PIN268' ,	'PIN270' ,	'PIN322' ,	'PIN323' ,	'PIN325' ,	'PIN326' ,	'PIN330' ,	'PIN331' ,	'PIN332' ,	'PIN334' ,	'PIN68'  ,	'PIN77'  ,	'PIN100' ,	'PIN109' ,	'PIN260' ,	'PIN269' ,	'PIN324' ,	'PIN333' ,	'PIN113' ,	'PIN114' ,	'PIN302' ,	'PIN303' ,	'PIN321' ,	'VS1'    ,	'VS2'    ,	'VS3']
PINLIST = ['CK_T_A'  ,	'CK_C_A'  ,	'CK_T_B'  ,	'CK_C_B'  ,	'CKE0_A'  ,	'CKE0_B'  ,	'CKE1_A'  ,	'CKE1_B'  ,	'CS0_A'   ,	'CS0_B'   ,	'CS1_A'   ,	'CS1_B'   ,	'CA0_A'   ,	'CA0_B'   ,	'CA1_A'   ,	'CA1_B'   ,	'CA2_A'   ,	'CA2_B'   ,	'CA3_A'   ,	'CA3_B'   ,	'CA4_A'   ,	'CA4_B'   ,	'CA5_A'   ,	'CA5_B'   ,	'ODT_CA_A',	'ODT_CA_B',	'DQ0_A'   ,	'DQ1_A'   ,	'DQ2_A'   ,	'DQ3_A'   ,	'DQ4_A'   ,	'DQ5_A'   ,	'DQ6_A'   ,	'DQ7_A'   ,	'DQ8_A'   ,	'DQ9_A'   ,	'DQ10_A'  ,	'DQ11_A'  ,	'DQ12_A'  ,	'DQ13_A'  ,	'DQ14_A'  ,	'DQ15_A'  ,	'DQ0_B'   ,	'DQ1_B'   ,	'DQ2_B'   ,	'DQ3_B'   ,	'DQ4_B'   ,	'DQ5_B'   ,	'DQ6_B'   ,	'DQ7_B'   ,	'DQ8_B'   ,	'DQ9_B'   ,	'DQ10_B'  ,	'DQ11_B'  ,	'DQ12_B'  ,	'DQ13_B'  ,	'DQ14_B'  ,	'DQ15_B'  ,	'DQS0_T_A',	'DQS0_C_A',	'DQS1_T_A',	'DQS1_C_A',	'DQS0_T_B',	'DQS0_C_B',	'DQS1_T_B',	'DQS1_C_B',	'DMI0_A'  ,	'DMI1_A'  ,	'DMI0_B'  ,	'DMI1_B'  ,	'RESET_N' ,	'VDD1'    ,	'VDD2'    ,	'VDDQ']
SHEETNAME = ['SUM', 'OS_H', 'OS_L', 'LEAKAGE', 'ODT', 'IDD', ]
CaseNumber = [100, 101, 102, 203, 503 ]

DATABANK = []
for q in range(700):
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
    #with open(file[0:-5]+'.xls', 'a') as f:         #测试代码，不起作用，勿动
    for i in range(len(content)):
        aaa = ('#' + content[i]).strip()
        if aaa.find('DUT') > 0 or aaa.find('TEST NAME') > 0 or aaa.find('#TNO') > 0:
            if aaa.find('UPPER') < 0 and aaa.find('[ps]') < 0 and aaa.find('Training Failed') < 0 and aaa.find('Fail to Pass') < 0 and aaa.find('RESULT') < 0 and aaa.find('Get Fail') < 0 and aaa.find('test sort') < 0:
                for idx in range(0, len(PINNUMB)):
                    aaa = aaa.replace(PINNUMB[idx], PINLIST[idx])
                data_array.append(aaa)
                #f.write(aaa+'\n')
    return data_array

def GetData(file, FILE_cunt, DATABANK, SUM_Array):     # excel关键信息筛选
    data_array = Select_data(file)          # 获取有效数据
    sumc = 0
    
    for q in range(700):           # 遍历所有case容器
        for i in range(4):          # 创建4个DUT数据容器
            DATABANK[q].append({})  # 创建4个DUT数据容器
    for i in range(4):
        SUM_Array.append([])        # 创建4个SUM数据容器

    for i in range(len(data_array)):                            # 遍历数据
        data = data_array[i]                                    # 方便下一步处理
        if data.find('#TNO') > 0:                               # SUM表内容提取
            datatemp = re.findall(r'Result:(.*?), Y', data)
            Ntemp = re.findall(r'Title:(.*?), R', data)
            Ttemp = re.findall(r'TT:(.*?), RT', data)
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
                SUM_Array[0][sumc].append(Ttemp[0])                # 获取测试时间    
            sumc += 1
        elif data.find('NAME:') > 0:
            data_name = re.findall(r'\d+', data)[0]
            dn = int(data_name)
        else:
            dvalue = data.split()[1]                     # 获取测量值

            if dn>=203 and dn<500:
                if dvalue.find('mA') > 0:
                    nup = re.findall(r'...\d+', dvalue)[0]
                    dvalue = 300 / float(nup)
                elif dvalue.find('uA') > 0:
                    nup = re.findall(r'\d+', dvalue)[0]
                    dvalue = 300 / ( float(nup) / 1000 )
                elif dvalue.find('A') > 0:
                    dvalue = 'NA'



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

    wsheet.write(0, 0, 'Test Case')                         # 打印tilte             |    
    wsheet.write(0, 1, 'Time')                              # 打印tilte             |    
    for num in range(len(SUM_Array[0])):                    #                       |打印SUM表
        wsheet.write(wrrow, 0, SUM_Array[0][num][0])        #                       |
        wsheet.write(wrrow, 1, SUM_Array[0][num][1])        #                       |
        wrrow += 1                                          #                       |
    for dut in range(1, len(SUM_Array)):                    #                       |
        wsheet.write(0, dut+1, 'DUT' + str(dut))            # 打印DUT编号            |
        wrrow = 1                                           #                       |
        for num in range(len(SUM_Array[dut])):              #                       |
            wsheet.write(wrrow, dut+1, SUM_Array[dut][num]) #                       |
            wrrow += 1                                      #                       |

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
    newfile = filepath +'SUM.xls'
    name = os.listdir(filepath)
    for i in name:
        print(i)
        RAW_file = filepath + i
        #Select_data(RAW_file)
        GetData(RAW_file, FILE_cunt, DATABANK, SUM_Array)
        FILE_cunt += 1
    PutDATA(newfile, DATABANK)
    end = time.perf_counter()
    print('程序执行时间: %.5s秒' % (end-start))
