# coding=utf-8
import os,subprocess
import threading
import socket
import pickle

path='C:\\Program Files\\Python37'
python_path=os.path.join(path,'python.exe')
excel_path=os.path.join(path,'excel.exe')

host,port='127.0.0.1',30000
#pool = redis.ConnectionPool(host=redis_host, port=6379)
#conn = redis.Redis(connection_pool=pool)

def main():
    App=Decr(top_dir='D:\\test\\C.CAD',size_limit=1024)
    App.run()
    #f=File('D:\\','7z1604-x64.exe')
    #f.tcp_decr(host,port)

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

    def run(self):
        server=subprocess.Popen([excel_path,'server.py'])
        self.get_files()
        for f in self.files:
            if f.size_m<self.size_limit:
                f.tcp_decr(host,port)
            else:
                print('over size: '+ f.fpath + str(f.size_m))
        server.kill()

class File():
    
    def __init__(self,path,name):
        self.path=path
        self.name=name
        self.fpath=os.path.join(path,name)
        self.size=os.path.getsize(self.fpath)
        self.size_m=self.size/1024/1024

    def load_data(self):
        self.data=open(self.fpath,'rb').read()
        return self.data

    def save(self,fpath=''):
        if not fpath:
            fpath=self.fpath
        try:
            with open(fpath,'wb') as f:
                f.write(self.data)
        except PermissionError:
            print('PermissionError: '+ self.fpath)
            

    def tcp_run(self):
        return self.load_data()

    def tcp_decr(self,host,port):
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # 建立连接:
        s.connect((host,port))
        # 接收欢迎消息:
        s.recv(1024)
        data=pickle.dumps(self)
        s.send(data)
        self.data=s.recv(self.size+1024**2)
        s.send('exit'.encode())
        s.close()
        self.save()
        self.data=''
        
def tcp_recv(host,port,obj):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    # 建立连接:
    s.connect((host,port))
    # 接收欢迎消息:
    s.recv(1024)
    data=pickle.dumps(obj)
    s.send(data)
    res=s.recv(obj.size+1024**2)
    s.send('exit'.encode())
    s.close()
    return res


if __name__ == '__main__':
    main()
