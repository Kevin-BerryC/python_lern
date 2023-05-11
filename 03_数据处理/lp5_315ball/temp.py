import time
import sys


def get_info(x):    #获取变量信息
    try:
        print(x,end=' ')
        print(type(x),len(x))
    except:
        print(type(x))
    else:
        return

def cycle():
    a = 'RECLIMIT:{}, EXEPATH:{}, UNICODE:{}'
    b = sys.getrecursionlimit(),sys.executable,sys.maxunicode
    #print(a.format(b))
    #print('RECLIMIT:{}, EXEPATH:{}, UNICODE:{}'.format(sys.getrecursionlimit(),sys.executable,sys.maxunicode))



if __name__ == '__main__':
    start = time.perf_counter()

    cycle()

    end = time.perf_counter()
    print('程序执行时间: %.5s秒' % (end-start))
