# -*- coding: utf-8 -*- 运行python
import rediscache

if __name__=="__main__":
    name='acadbindhelper'
    rediscache.download(rediscache.conn,name)
    import acadbindhelper as helper
    helper.method=1
    helper.autorun()

