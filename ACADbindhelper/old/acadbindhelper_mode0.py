# -*- coding: utf-8 -*- 运行python
import acadbindhelper as helper
import warnings

if __name__=="__main__":
    warnings.simplefilter("once")
    helper.method=0
    helper.autorun()
    while 1:
        try:
            helper.autorun()
        except:
            pass


