import xlsxwriter
import time
import re
import os
import sys

if __name__ == '__main__':
    start = time.perf_counter()
    filepath = 'D:\\DATA\\INPUT\\'
    name = os.listdir(filepath)

    '''#for i in range(len(name)):
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
            PutDATA(newfile, DATABANK, INFOBANK)'''
    end = time.perf_counter()
    print('程序执行时间: %.5s秒' % (end-start))

