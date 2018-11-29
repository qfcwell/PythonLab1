# -*- coding: utf-8 -*-
from flask import Flask
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'mssql+pymssql://sa:qfc23834358Q@10.1.42.131/Product_Test'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = True
db = SQLAlchemy(app)

class P_Class(db.Model):#产品类，族
    __tablename__ = 'P_Class'
    Id=db.Column(db.Integer, primary_key=True)
    ClassName=db.Column(db.String(50))
    ClassType=db.Column(db.String(50))
    PicURL=db.Column(db.String(500))
    Table=db.Column(db.String(50))
    NoFill=db.Column(db.String(50))
    MustFill=db.Column(db.String(50))
    TypePara=db.Column(db.String(50))
    EntPara=db.Column(db.String(50))
    IsActive=db.Column(db.Integer,default=1)
    Paras=db.relationship('P_Para', backref='P_Class',lazy='dynamic')

    def __repr__(self):
        return '<P_Class %r>' % self.ClassName

class P_Para(db.Model):#参数
    __tablename__ = 'P_Para'
    Id=db.Column(db.Integer, primary_key=True)
    ClassId=db.Column(db.Integer, db.ForeignKey('P_Class.Id'))
    ParaName=db.Column(db.String(50))
    FieldName=db.Column(db.String(50))
    Metric=db.Column(db.String(50))
    Unit=db.Column(db.String(50))
    RFieldName=db.Column(db.String(50))
    IsActive=db.Column(db.Integer,default=1)

    def __repr__(self):
        return '<P_Para %r>' % self.ParaName

P_Class_Category=db.Table("P_Class_Category",
    db.Column('P_Class_Id',db.Integer,db.ForeignKey('P_Class.Id')),
    db.Column('P_Category_Id',db.Integer,db.ForeignKey('P_Category.Id'))
    )

class P_Category(db.Model):#产品类分类目录
    __tablename__ = 'P_Category'
    Id=db.Column(db.Integer, primary_key=True)
    CategoryName=db.Column(db.String(50))
    LogicalPath=db.Column(db.String(500))
    RealPath=db.Column(db.String(500))
    IsActive=db.Column(db.Integer,default=1)
    ChildClasses=db.relationship('P_Class',secondary=P_Class_Category,backref=db.backref('Categories',lazy='dynamic'))

    def __repr__(self):
        return '<P_Category %r>' % self.CategoryName



if __name__=="__main__":
    db.drop_all()
    db.create_all()
    #session.add(P_Class(Name='111'))
    #session.commit()
