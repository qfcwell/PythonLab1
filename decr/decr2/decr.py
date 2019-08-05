# coding=utf-8
import os,subprocess
import redis
import time,uuid

def main():
    App=Decr(top_dir='D:\\decr')
    App.decr()
    #_file("D:\\decr\\2017\\二维协同模块化设计深化计划书(2016.12.13).docx")

class Decr():
    def __init__(self,top_dir,size_limit=1024,redis_host='superset1.minitech.site',key='decr_cache',path='C:\\Program Files\\Python35'):
        self.pool = redis.ConnectionPool(host=redis_host, port=6379)
        self.conn = redis.Redis(connection_pool=self.pool)
        self.path=path
        self.python_path=os.path.join(path,'python.exe')
        self.excel_path=os.path.join(path,'excel.exe')
        self.f_list=[]
        self.files=[]
        self.top_dir=top_dir
        self.redis_host=redis_host
        self.key=key
        self.size_limit=size_limit

    def create_excel_exe(self):
        with open(self.python_path,'rb') as f1:
            with open(self.excel_path,'wb') as f2:
                f2.write(f1.read())

    def remove_excel_exe(self):
        os.remove(self.excel_path)

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
                    self.f_list.append(os.path.join(dir_path,name))
                    self.files.append(File(dir_path,name))
        return self.f_list

    def upload(self,fpath):
        sp=subprocess.call([self.excel_path,'upload.py',self.redis_host,fpath,self.key])

    def decr_file(self,fpath):
        sp=subprocess.call([self.excel_path,'upload.py',self.redis_host,fpath,self.key])
        res=self.conn.get(self.key)
        with open(fpath,'wb') as f:
            f.write(res)

    def decr(self):
        self.create_excel_exe()
        self.get_files()
        for fpath in self.f_list:
            if os.path.getsize(fpath)/1024/1024>self.size_limit:
                print('oversize: '+fpath)
            else:
                self.upload(fpath)
                res=self.conn.get(self.key)
                if res:
                    try:
                        with open(fpath,'wb') as f:
                            f.write(res)
                            print('decr '+fpath)
                    except PermissionError:
                        print('PermissionError on '+fpath)
                else:
                    print(fpath+' is none,size: '+str(os.path.getsize(fpath)/1024/1024))
                self.conn.delete(self.key)
        self.remove_excel_exe()

class File():
    
    def __init__(self,path,name,redis_host='superset1.minitech.site'):
        self.redis_host=redis_host
        self.pool = redis.ConnectionPool(host=redis_host, port=6379)
        self.conn = redis.Redis(connection_pool=self.pool)
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
        sp=subprocess.call([self.excel_path,'upload.py',self.redis_host,self.fpath,self.key])

    def download(self):
        pass


if __name__ == '__main__':
    main()