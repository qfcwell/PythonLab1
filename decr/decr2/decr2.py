# coding=utf-8
import os,subprocess
import redis
import time,uuid
import threading
from multiprocessing import Pool
from multiprocessing.dummy import Pool as ThreadPool

path='C:\\Program Files\\Python35'
python_path=os.path.join(path,'python.exe')
excel_path=os.path.join(path,'excel.exe')
redis_host='superset1.minitech.site'
pool = redis.ConnectionPool(host=redis_host, port=6379)
conn = redis.Redis(connection_pool=pool)

def main():
    App=Decr(top_dir='D:\\decr')
    App.run(m=1)
    #_file("D:\\decr\\2017\\二维协同模块化设计深化计划书(2016.12.13).docx")

class Decr():
    def __init__(self,top_dir,size_limit=1024):
        self.files=[]
        self.top_dir=top_dir
        self.size_limit=size_limit

    def get_files(self):
        top_dir=self.top_dir
        if os.path.exists(top_dir)==False:
            print("not exists")
            return 0
        elif os.path.isdir(top_dir)==False:
            print("not a dir")
            return 0
        else:
            for dir_path,subpaths,files in os.walk(top_dir,False):
                for name in files:
                    #self.f_list.append(os.path.join(dir_path,name))
                    self.files.append(File(dir_path,name))
        return self.files

    def run(self,m=0):
        self.get_files()
        if m:
            #pool = ThreadPool(100) # Sets the pool size to 4
            for f in self.files:
                #decr_file(f)
                t=threading.Thread(target=decr_file,args=(f,))
                t.start()
                #t=pool.map(decr_file, (f,))
            #pool.close() 
        else:
            for f in self.files:
                decr_file(f)

class File():
    
    def __init__(self,path,name):
        self.path=path
        self.name=name
        self.fpath=os.path.join(path,name)
        self.uid=str(uuid.uuid1())
        self.size=os.path.getsize(self.fpath)

    def rename_uid(self):
        cfpath=os.path.join(self.path,self.uid)
        os.rename(self.fpath,cfpath)
        return cfpath

    def rename_name(self):
        cfpath=os.path.join(self.path,self.uid)
        os.rename(cfpath,self.fpath)
        return fpath

    def upload(self):
        sp=subprocess.call([excel_path,'upload.py',redis_host,self.fpath,self.uid])

    def download(self):
        res=conn.get(self.uid)
        try:
            with open(self.fpath,'wb') as f:
                f.write(res)
        except PermissionError:
            print('PermissionError on '+self.fpath)
        conn.delete(self.uid)

    def decr(self):
        if self.size/1024/1024>1024:
            print('oversize: '+self.fpath)
        else:
            print('decr ' + self.fpath)
            self.upload()
            self.download()

def decr_file(file):
    if file.size/1024/1024>1024:
        print('oversize: '+file.fpath)
    else:
        print('decr ' + file.fpath)
        file.upload()
        file.download()

if __name__ == '__main__':
    main()