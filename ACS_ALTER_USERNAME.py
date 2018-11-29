# -*- coding: utf-8 -*-
import pymssql
'''
def run():
    with pymssql.connect(host='10.1.246.1', user="sa", password="sasasa",database="CAPOL_Project") as conn:
        cur=conn.cursor()
        UserId=GetUserId(cur,'libina')
        print(UserId)
        AlterUserName(cur,UserId,'李斌2')
        AlterLoginName(cur,UserId,'libin2')
        conn.commit()
'''
def run():
    with pymssql.connect(host='10.1.246.1', user="sa", password="sasasa",database="CAPOL_Project") as conn:
    #with pymssql.connect(host='10.2.1.114', user="sa", password="hygj!@#456",database="CAPOL_Project") as conn:
        cur=conn.cursor()
        UserId=GetUserId(cur,'jgjs1')
        print(UserId)
        AlterUserName(cur,UserId,'结审1')
        AlterLoginName(cur,UserId,'jieshen1')
        conn.commit()

def GetUserId(cur,LoginName):
    cur.execute(u"SELECT ID,UserName FROM CAPOL_Core.dbo.SysUser where LoginName=%s",(LoginName))
    (UserId,UserName)=cur.fetchone()
    return UserId

def AlterUserName(cur,UserId,NewName):#改中文姓名
    cur.execute(u"SELECT UserName FROM CAPOL_Core.dbo.SysUser where ID=%s",(UserId))
    (OldName,)=cur.fetchone()
    if OldName==NewName:
        return 0
    else:
        cur.execute(u"UPDATE CAPOL_Core.dbo.SysUser set UserName=%s where ID=%s",(NewName,UserId))
        cur.execute(u"UPDATE DesignFile set CreatorName=%s where CreatorId=%s",(NewName,UserId))#刷设计文件
        cur.execute(u"UPDATE DesignFile set LastBackuperName=%s where LastBackuperId=%s",(NewName,UserId))#刷设计文件
        cur.execute(u"UPDATE DesignFile set LockerName=%s where LockerId=%s",(NewName,UserId))#刷设计文件
        cur.execute(u"UPDATE TakeoutFile set TakeoutUserName=%s where TakeoutUserId=%s",(NewName,UserId))#刷提资文件
        cur.execute(u"UPDATE ProductFile set CreatorName=%s where CreatorId=%s",(NewName,UserId))#刷归档文件
        cur.execute(u"SELECT Id,ExecutorId,ExecutorName from PrjRole WHERE ExecutorId is not null and ExecutorId!=''")
        for (RoleId,ExecutorId,ExecutorName) in cur.fetchall():
            if UserId in ExecutorId.split(','):
                lst=[]
                for uid in ExecutorId.split(','):
                    cur.execute(u"SELECT UserName FROM CAPOL_Core.dbo.SysUser WHERE Id=%s",(uid))
                    (uname,)=cur.fetchone()
                    lst.append(uname)
                NewExecutorName=','.join(lst)
                cur.execute(u"UPDATE PrjRole SET ExecutorName=%s where Id=%s",(NewExecutorName,RoleId))

def AlterLoginName(cur,UserId,NewName):
    cur.execute(u"UPDATE CAPOL_Core.dbo.SysUser set LoginName=%s where ID=%s",(NewName,UserId))#修改用户名


if __name__=="__main__":
    run()
