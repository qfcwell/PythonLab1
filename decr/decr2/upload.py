import redis,sys

def main():
    upload(sys.argv[1],sys.argv[2],sys.argv[3])
    
def upload(redis_host,fpath,key):
    pool = redis.ConnectionPool(host=redis_host, port=6379)
    conn = redis.Redis(connection_pool=pool)
    with open(fpath,'rb') as f:
        res=f.read()
    conn.set(key,res)
    return 0

def download(redis_host,fpath,key):
    pool = redis.ConnectionPool(host=redis_host, port=6379)
    conn = redis.Redis(connection_pool=pool)
    res=conn.get(key)
    with open(fpath,'wb') as f:
        f.write(res)
    return 0


if __name__ == '__main__':
    main()