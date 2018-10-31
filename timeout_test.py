# coding:utf-8
import time
from threading import Timer

class TimeOutError(Exception):
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr(self.value)

def timeout(interval):
    def wraps(func):
        def time_out():
            raise TimeOutError('TimeOut: '+str(interval))
        def deco(*args, **kwargs):
            timer = Timer(interval, time_out)
            timer.start()
            res = func(*args, **kwargs)
            timer.cancel()
            return res
        return deco
    return wraps

@timeout(5)
def mytest():
    print("Start")
    for i in range(1,10):
        time.sleep(1)
        print("%d seconds have passed" % i)

if __name__ == '__main__':
    mytest()