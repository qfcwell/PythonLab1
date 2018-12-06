# -*- coding: utf-8 -*-
import pymssql,cx_Oracle
import os


os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
acs_server_lst={'shenzhen':'10.1.246.1','guangzhou':'10.2.1.114','changsha':'10.3.1.3','shanghai':'10.6.1.15','hainan':'10.14.2.10'}

def run():
    #init_cache()
    with InOutSideMark_Check(method='capol') as check:
        check.RunCheck()

def init_cache():
    for S_NAME in acs_server_lst:
        conn=pymssql.connect(host='10.1.42.131', user="sa", password="qfc23834358Q",database="CADtest")
        cur=conn.cursor()
        sql="""SELECT '%(S_NAME)s' as S_NAME,TR.SubEntryId,TR.TakeoutMajorCode,TRF.TakeoutFloorName,TR.PutinMajorCode,TRF.PutinFloorName 
            INTO CAPOL_TRF_CACHE_%(S_NAME)s
            FROM [%(link_server)s].CAPOL_PROJECT.DBO.TakeoutRelation TR 
            JOIN [%(link_server)s].CAPOL_PROJECT.DBO.TakeoutRelationFloor TRF on TR.ID=TRF.TakeoutRelationId
            WHERE TRF.TakeoutFloorName IS NOT NULL AND TRF.PutinFloorName IS NOT NULL""" % {'S_NAME':S_NAME,'link_server':acs_server_lst[S_NAME]}
        cur.execute(sql)
        conn.commit()

class sync():
    def __init__(self,method='test'): 
        self.cip=pymssql.connect(host='10.1.1.117', user="cip_ro", password="hygj!@#4567",database="cip")
        self.oracle_ccd=cx_Oracle.connect('ccd/hygj1234@10.1.1.213:1521/orcl')
        self.oracle_test=cx_Oracle.connect('steven/qfc23834358@10.1.42.131:1521/xe')
        if method=='test':
            self.oracle=self.oracle_test
        else:
            self.oracle=self.oracle_ccd
        self.acs_server={}
        self.acs_server_lst=acs_server_lst
        for S_NAME in acs_server_lst:
            self.acs_server[S_NAME]=pymssql.connect(host=acs_server_lst[S_NAME], user="readonly", password="capol!@#456",database="CAPOL_Project")

    def __enter__(self):
        return self

    def __exit__(self, type, value, trace):
        self.close()

    def close(self):
        self.a=1
        self.cip.close()
        self.oracle_ccd.close()
        self.oracle_test.close()
        for S_NAME in self.acs_server:
            self.acs_server[S_NAME].close()

class InOutSideMark_Check(sync):

    def RunCheck(self):
        for S_NAME in acs_server_lst:
            with DatumnRelationCache(S_NAME) as DRC:
                Change=DRC.Change
                DRC.Update(S_NAME)
                print(Change)
            for (S_NAME,SubEntryId,TakeoutMajorCode,TakeoutFloorName,PutinMajorCode,PutinFloorName) in Change:
                OutFiles=self.FindTakeoutFile(S_NAME,SubEntryId,TakeoutMajorCode,TakeoutFloorName)
                InFiles=self.FindTakeinFile(S_NAME,SubEntryId,PutinMajorCode,PutinFloorName,FileTypeCode='SingleFrame')
                self.UpdateInOutSideMarkRelation(S_NAME,OutFiles,PutinMajorCode,InFiles)
            for (SubEntryId,MajorCode,Floor,TakeOutTime) in self.FindDatumnEvent(S_NAME):
                OutFiles=self.FindTakeoutFile(S_NAME,SubEntryId,MajorCode,Floor)
                self.UpdateInOutSideMarkState(S_NAME,OutFiles,TakeOutTime)

    def SubEntryList(self,S_NAME):
        lst=[]
        cur=self.acs_server[S_NAME].cursor()
        cur.execute(u"""SELECT ID FROM SubEntry WHERE RecordState='A'""")
        for (SubEntryId,) in cur.fetchall():
            lst.append(SubEntryId)
        return lst

    #查询近30分钟提资的项目阶段子项，返回子项ID、提出专业、提出楼层、提资时间
    def FindDatumnEvent(self,S_NAME):
        cur=self.acs_server[S_NAME].cursor()
        cur.execute(u"""SELECT SubEntryId,MajorCode,Floor,convert(varchar(120),TakeoutTime,20) FROM TakeoutFloorRecord WHERE TakeoutTime>dateadd(mi,-30,getdate()) """)
        res=cur.fetchall()
        return res

    #查询有变化的提资匹配关系，返回子项ID、提出专业、提出楼层、接收专业、接收楼层
    def FindDatumnRelationChange(self,S_NAME):
        pass
        

    #查询提出文件，输入协同服务器名、子项ID、提出专业、提出楼层，返回相关文件全路径列表
    def FindTakeoutFile(self,S_NAME,SubEntryId,Major,Floor):
        lst=[]
        cur=self.acs_server[S_NAME].cursor()
        cur.execute(u"""SELECT Project.FolderName,PrjPhase.FolderName,SubEntry.FolderName 
            FROM Project JOIN PrjPhase ON PrjPhase.PrjId=Project.Id 
            JOIN SubEntry ON SubEntry.PrjPhaseId=PrjPhase.Id 
            WHERE SubEntry.Id=%s AND SubEntry.RecordState='A' """,(SubEntryId,))
        res=cur.fetchone()
        if res:
            (prj,phase,sub)=res
            path=os.path.join('D:\\AcsModule\\WorkFolder\\Project',prj+'_'+phase,sub)
            cur.execute(u"""SELECT SysFileName FROM DesignFile where SubEntryId=%s AND MajorCode=%s AND Floors=%s """,(SubEntryId,Major,Floor))
            res=cur.fetchall()
            for (f,) in res:
                fpath=os.path.join(path,'设计区',f)
                lst.append(fpath)
        return lst

    #查询接收文件，输入协同服务器名、子项ID、接收专业、接收楼层，返回框架文件列表
    def FindTakeinFile(self,S_NAME,SubEntryId,Major,Floor,FileTypeCode=''):
        lst=[]
        cur=self.acs_server[S_NAME].cursor()
        if FileTypeCode:
            cur.execute(u"""SELECT SysFileName FROM DesignFile where SubEntryId=%s AND MajorCode=%s AND Floors=%s AND FileTypeCode=%s""",(SubEntryId,Major,Floor,FileTypeCode))
        else:
            cur.execute(u"""SELECT SysFileName FROM DesignFile where SubEntryId=%s AND MajorCode=%s AND Floors=%s """,(SubEntryId,Major,Floor))
        res=cur.fetchall()
        for (f,) in res:
            lst.append(f)
        return lst

    #更新专业间意见的匹配关系，输入协同服务器名、文件全路径、接收专业、接收框架文件名
    def UpdateInOutSideMarkRelation(self,S_NAME,OutFiles,InMajor,InFiles):
        cur=self.oracle.cursor()
        sql1="'"+";".join(InFiles)+"'"
        table_dic={"A":"ARCHFILENAME","S":"STRUFILENAME","M":"MECHFILENAME","W":"PUMPFILENAME","E":"ELECFILENAME","V":"TELEFILENAME","G":"GASFILENAME"}
        cur.execute(u"SELECT 文件路径 FROM capol_outsidemark")
        NotOutFiles=cur.fetchall()
        for outfile in OutFiles:
            if (outfile,) in NotOutFiles:
                sql="UPDATE capol_outsidemark SET %(table)s=%(infile)s WHERE 文件路径 = '%(outfile)s'" % {"table":table_dic[InMajor],"infile":sql1,"outfile":outfile}
                print(sql)
                cur.execute(sql)
        self.oracle.commit()

    #更新专业间意见的状态，输入协同服务器名、文件全路径、提资时间，如果提资时间大于标记添加时间且标记状态为未提资，更新为已提资
    def UpdateInOutSideMarkState(self,S_NAME,OutFiles,TakeOutTime):
        cur=self.oracle.cursor()
        TakeOutTime="'"+str(TakeOutTime)+"'"
        cur.execute(u"SELECT 文件路径 FROM capol_outsidemark WHERE ISPUTOUT=0")
        NotOutFiles=cur.fetchall()
        for outfile in OutFiles:
            if (outfile,) in NotOutFiles:
                sql="UPDATE capol_outsidemark SET ISPUTOUT=1 WHERE 添加时间<to_date(%(TakeOutTime)s,'YYYY-MM-DD HH24:MI:SS') AND ISPUTOUT=0 AND 文件路径 = %(outfile)s" % {"outfile":outfile,"TakeOutTime":TakeOutTime}
                print(sql)
                cur.execute(sql)
        self.oracle.commit()

class DatumnRelationCache(sync):
    def __init__(self, S_NAME, *arg, **kw):
        self.acs_server_lst=acs_server_lst
        self.acs=pymssql.connect(host=self.acs_server_lst[S_NAME], user="readonly", password="capol!@#456",database="CAPOL_Project")
        self.cache_server=pymssql.connect(host='10.1.42.131', user="sa", password="qfc23834358Q",database="CADtest")
        self.S_NAME=S_NAME
        self.Change=self.Get_Change(S_NAME)

    def close(self):
        self.acs.close()
        self.cache_server.close()

    def Get_Change(self,S_NAME):
        cur=self.cache_server.cursor()
        sql="""SELECT '%(S_NAME)s' as S_NAME,TR.SubEntryId,TR.TakeoutMajorCode,TRF.TakeoutFloorName,TR.PutinMajorCode,TRF.PutinFloorName 
            FROM [%(link_server)s].CAPOL_PROJECT.DBO.TakeoutRelation TR 
            JOIN [%(link_server)s].CAPOL_PROJECT.DBO.TakeoutRelationFloor TRF on TR.ID=TRF.TakeoutRelationId
            WHERE TRF.TakeoutFloorName IS NOT NULL AND TRF.PutinFloorName IS NOT NULL
            except 
            SELECT S_NAME,SubEntryId,TakeoutMajorCode,TakeoutFloorName,PutinMajorCode,PutinFloorName
            FROM CAPOL_TRF_CACHE_%(S_NAME)s""" % {'S_NAME':S_NAME,'link_server':acs_server_lst[S_NAME]}
        print(sql)
        cur.execute(sql)
        res=cur.fetchall()
        return res

    def Update(self,S_NAME):
        cur=self.cache_server.cursor()
        cur.execute(u"""DROP TABLE CAPOL_TRF_CACHE_%(S_NAME)s""" % {'S_NAME':S_NAME})
        sql="""SELECT '%(S_NAME)s' as S_NAME,TR.SubEntryId,TR.TakeoutMajorCode,TRF.TakeoutFloorName,TR.PutinMajorCode,TRF.PutinFloorName 
            INTO CAPOL_TRF_CACHE_%(S_NAME)s
            FROM [%(link_server)s].CAPOL_PROJECT.DBO.TakeoutRelation TR 
            JOIN [%(link_server)s].CAPOL_PROJECT.DBO.TakeoutRelationFloor TRF on TR.ID=TRF.TakeoutRelationId
            WHERE TRF.TakeoutFloorName IS NOT NULL AND TRF.PutinFloorName IS NOT NULL""" % {'S_NAME':S_NAME,'link_server':acs_server_lst[S_NAME]}
        print(sql)
        cur.execute(sql)
        self.cache_server.commit()
        return True



if __name__=="__main__":
    run()
