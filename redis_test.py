# -*- coding: utf-8 -*- 运行python
import redis,os
pool = redis.ConnectionPool(host='10.1.42.132', port=6379)
#pool = redis.ConnectionPool(host='39.108.109.58', port=6379)
#pool = redis.ConnectionPool(host='10.1.246.1', port=6379)
conn = redis.Redis(connection_pool=pool)

def main():
    keys = conn.keys()
    for k in keys:
        print(k)

def get_all_file(path='.\\',not_in=[]):
    lst=[]
    dirs=os.listdir(path)
    for d in dirs:
        if os.path.isfile(d) and d not in not_in:
            lst.append(d)
    return lst

def upload(conn,name,lst,delete=False):
    if delete:
        conn.delete(name)
    for file in lst:
        with open(file,'rb') as f:
            res=f.read()
            conn.hset(name,file,res)
            print('update:'+file)

def download(conn,name):
    keys=conn.hkeys(name)
    for key in keys:
        with open(key,'wb') as f:
            f.write(conn.hget(name,key))

if __name__ == '__main__':
    main()