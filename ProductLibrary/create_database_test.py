# -*- coding: utf-8 -*-
from flask import Flask
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'mssql+pymssql://sa:qfc23834358Q@10.1.42.131/Product_Test'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = True
db = SQLAlchemy(app)

P_Class_Category = db.Table('P_Class_Category',
    db.Column('P_Class_Id', db.Integer, db.ForeignKey('P_Class.Id')),
    db.Column('P_Category_Id', db.Integer, db.ForeignKey('P_Category.Id'))
)

class P_Category(db.Model):
    __tablename__ = 'P_Category'
    Id = db.Column(db.Integer, primary_key=True)
    P_Class_Category = db.relationship('P_Class', secondary=P_Class_Category,
        backref=db.backref('P_Categorys', lazy='dynamic'))

class P_Class(db.Model):
    Id = db.Column(db.Integer, primary_key=True)


if __name__=="__main__":
    #db.drop_all()
    db.create_all()
    #session.add(P_Class(Name='111'))
    #session.commit()
