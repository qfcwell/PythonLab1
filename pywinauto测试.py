# -*- coding: utf-8 -*- 运行python
import os
from pywinauto.application import Application

def run():
    app=Application().connect(title_re="提示",class_name="Dialog")

if __name__=="__main__":
    run()

