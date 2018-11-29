
from sqlalchemy import create_engine, Table
from sqlalchemy import Column, ForeignKey
from sqlalchemy import String, Integer, Text
from sqlalchemy.orm import sessionmaker, relationship
from sqlalchemy.ext.declarative import declarative_base

Base = declarative_base()
engine = create_engine('mssql+pymssql://sa:qfc23834358Q@10.1.42.131/Product_Test')
Session = sessionmaker(bind=engine)
session = Session()

class P_Class(Base):#产品类，族
    __tablename__='P_Class'
    Id=Column(Integer, primary_key=True)
    Name=Column(String(50))
    Type=Column(String(50))
    PicURL=Column(String(500))
    Table=Column(String(50))
    NoFill=Column(String(50))
    MustFill=Column(String(50))
    TypePara=Column(String(50))
    EntPara=Column(String(50))
    ParaList=relationship('P_Para')

    def __repr__(self):
        return '%s(%r)' % (self.__class__.__name__, self.Name)

class P_Para(Base):#参数
    __tablename__='P_Para'
    Id=Column(Integer, primary_key=True)
    P_Class_Id=Column(Integer, ForeignKey('P_Class.Id'))
    Name=Column(String(50))
    FieldName=Column(String(50))
    Metric=Column(String(50))
    Unit=Column(String(50))
    RFieldName=Column(String(50))

    def __repr__(self):
        return '%s(%r)' % (self.__class__.__name__, self.Name)




if __name__=="__main__":
    Base.metadata.create_all(engine)
    session.add(P_Class(Name='111'))
    session.commit()
