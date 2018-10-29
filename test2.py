# -*- coding: utf-8 -*-
import pymssql,cx_Oracle
import os



def run():
    for res in fib2(10000):
        print(res)

def fib(num):
    if num<2:        
        return num     
    return fib(num-1)+fib(num-2)

def fib2(max):
    a=b=n=1
    while n<max:
        yield a
        a,b=b,a+b
        n=n+1

if __name__=="__main__":
    run()
