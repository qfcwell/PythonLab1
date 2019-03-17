#! /usr/bin/env python   
# -*- coding: utf-8 -*-   

import os 
import time 
import uuid  
import pymysql

sourceDir = "D:\\decr" 
tempDir = "S:\\部门文件\\集团公司\\运营管理中心\\其他\\decr"  
targetDir = "D:\\decr_d"
#copyFileCounts = 0 
#F_list=[]
  
def copyFiles(sourceDir, targetDir):   
    copyFileCounts = 0 
    F_list=[]
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
                F_list.append(targetF)
                print(u"%s %s 复制完毕" %(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())), targetF))
            else:   
                print(u"%s %s 已存在，不重复复制" %(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())), targetF))
        if os.path.isdir(sourceF):   
            copyFiles(sourceF, targetF)
    return F_list

def newName(oldName):
    uid=uuid.uuid1()
    with pymysql.connect(host='10.1.42.103',user='steven',passwd='steven!@#456',db='project_test') as cur:
        cur.execute('INSERT INTO uuid_rename(uuid,name) VALUES(%s,%s)',[uid,oldName])

class File():
    def __init__(self,path,name):
        self.path=path
        self.name=name
        self.fpath=os.path.join(path,name)
        self.uid=str(uuid.uuid1())
        self.png=self.uid+'.png'
        self.xls=self.uid+'.xls'

    def rename_uid(self):
        os.rename

    def log_uid(self):
        with pymysql.connect(host='10.1.42.103',user='steven',passwd='steven!@#456',db='project_test') as cur:
            cur.execute('INSERT INTO uuid_rename(uuid,name) VALUES(%s,%s)',[self.uid,self.fpath])
    
    def rename_xls(self):
        cfpath=os.path.join(self.path,self.xls)
        os.rename(self.fpath,cfpath)
        return cfpath

    def rename_png(self):
        cfpath=os.path.join(tempDir,self.png)
        os.rename(os.path.join(tempDir,self.xls),cfpath)
        return cfpath

    def copy_temp(self):
        targetF=os.path.join(tempDir,self.xls)
        sourceF=os.path.join(self.path,self.xls)
        c='"D:\\GitHub\\decr\\dist\\excel.exe" "'+sourceF+'" "'  + targetF + '"'
        res=os.popen(c)
        res.close()
        cfpath=targetF
        return cfpath

    def copy_back(self):
        sourceF=os.path.join(tempDir,self.png)
        targetF=self.fpath
        try:
            with open(targetF, "wb") as f:
                f.write(open(sourceF, "rb").read())
            return self.fpath
        except PermissionError:
            print(self.fpath)
            return 0

    def del_temp(self):
        targetF=os.path.join(self.path,self.xls)
        os.remove(targetF)
        targetF=os.path.join(tempDir,self.png)
        os.remove(targetF)

def get_Files(top_dir):
    if os.path.exists(top_dir)==False:
        print("not exists")
        return 0
    elif os.path.isdir(top_dir)==False:
        print("not a dir")
        return 0
    else:
        lst=[]
        for dir_path,subpaths,files in os.walk(top_dir,False):
            for name in files:
                f=File(dir_path,name)
                lst.append(f)
        for f in lst:
            #f.log_uid()
            print('copy temp: '+ f.name)
            f.rename_xls()
            f.copy_temp()
        print("finish copy...")
        time.sleep(3)
        for f in lst:
            print('copy back: '+ f.name)
            f.rename_png()
            f.copy_back()
            f.del_temp()


if __name__ == "__main__": 
    get_Files(top_dir=sourceDir)
    #print(f)
    '''
    copyFiles(sourceDir,tempDir)
    for file in F_list:
        os.rename(file,file+".png")'''