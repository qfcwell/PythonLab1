# -*- coding: utf-8 -*- 运行python
import acadbindhelper as helper

if __name__=="__main__":
    helper.method=0
    helper.autorun()
    while 1:
        try:
            helper.autorun()
        except:
            pass


