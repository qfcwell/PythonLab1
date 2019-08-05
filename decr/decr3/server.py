# coding=utf-8

import threading
import socket
import pickle
from decr import File

host,port='127.0.0.1',30000

def main():
    serv = threading.Thread(target=server, args=(host,port))
    serv.start()

def server(host,port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind((host,port))
    s.listen(5)
    print('waiting for connection...')
    while True:
        # 接受一个新连接:
        sock, addr = s.accept()
        # 创建新线程来处理TCP连接:
        t = threading.Thread(target=tcp_recv, args=(sock, addr))
        t.start()

def tcp_recv(sock, addr):
    print('Accept new connection from %s:%s...' % addr)
    sock.send('hello!'.encode())
    max_len=1024
    while 1:
        data = sock.recv(max_len)
        if data == 'exit'.encode() or not data:
            break
        #try:
        obj=pickle.loads(data)
        res=obj.tcp_run()
        sock.send(res)
        #except:
         #   break
    sock.close()
    print('Connection from %s:%s closed.' % addr)


if __name__ == '__main__':
    main()
