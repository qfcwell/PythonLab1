# -*- coding: utf-8 -*- 运行python
from create_database2 import *
import os

root_path='D:\\01-标准族库'

def main():
    pass

def Find_Categories(root_path):
    for (fpath,children,cl) in os.walk(root_path):
        (ppath,cname)=os.path.split(fpath)
        Father=P_Category.query.filter_by(RealPath=ppath).first()
        if Father:
            LogicalPath=Father.LogicalPath+'/'+cname
        else:
            LogicalPath=''
        db.session.add(P_Category(CategoryName=cname,RealPath=fpath,LogicalPath=LogicalPath))
        ClassId=P_Category.query.filter_by(RealPath=fpath).first().Id
        for ClassName in cl:


    db.session.commit()

def Create_Category(CategoryName,SubCategoryId):



if __name__=="__main__":
    db.drop_all()
    db.create_all()
    #P_Category.query.filter_by(Path=ppath).all()
    #P_Category.query.filter().all()
    #db.session.query(P_Category).filter_by(Path=root_path).all()
    #Find_Categories(root_path)
