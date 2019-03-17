#! /usr/bin/env python   
# -*- coding: utf-8 -*-   
import sys

def main():
    copy_temp(sys.argv[1],sys.argv[2])

def copy_temp(sourceF,targetF):
    with open(targetF, "wb") as fw:
        with open(sourceF, "rb") as fr: 
            fw.write(fr.read())

if __name__ == '__main__':
    main()