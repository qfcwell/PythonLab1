# coding=utf-8
import redis


if __name__ == '__main__':
    pool = redis.ConnectionPool(host='superset1.minitech.site', port=6379)
    conn = redis.Redis(connection_pool=pool)

    keys=conn.keys()
    for key in keys:
        print(key)
        conn.delete(key)
