#! /usr/bin/env python   
# -*- coding: utf-8 -*-   

import os 
import time   

sourceDir = "D:\\GitHub" 
targetDir = "D:\\GitHub_c"  
copyFileCounts = 0 
  
def copyFiles(sourceDir, targetDir):   
    global copyFileCounts 
    print(sourceDir)
    print(u"%s 当前处理文件夹%s已处理%s 个文件" %(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())), sourceDir,copyFileCounts))
    for f in os.listdir(sourceDir):  
        sourceF = os.path.join(sourceDir, f)  
        targetF = os.path.join(targetDir, f) 
        if os.path.isfile(sourceF):   
            #创建目录   
            if not os.path.exists(targetDir): 
                os.makedirs(targetDir)  
            copyFileCounts += 1 
            #文件不存在，或者存在但是大小不同，覆盖   
            if not os.path.exists(targetF) or (os.path.exists(targetF) and (os.path.getsize(targetF) != os.path.getsize(sourceF))): 
                #2进制文件  
                open(targetF, "wb").write(open(sourceF, "rb").read()) 
                print(u"%s %s 复制完毕" %(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())), targetF))
            else:   # ~# k1 b! N) [% u
                print(u"%s %s 已存在，不重复复制" %(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())), targetF))
           
        if os.path.isdir(sourceF):   
            copyFiles(sourceF, targetF)   

if __name__ == "__main__": 
    try:   
        import psyco 
        psyco.profile()   
    except ImportError: 
        pass 
    copyFiles(sourceDir,targetDir)