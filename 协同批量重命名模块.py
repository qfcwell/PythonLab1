import pymssql


##替换文件名称

def main():
    conn = pymssql.connect(host="10.1.246.1", user="sa", password="sasasa",database="CAPOL_Project")  
    cur = conn.cursor()
    sql="SELECT Id,SysFileName FROM DesignFile WHERE PrjPhaseId='PPS20BKF' AND MajorCode='S' AND DrawingTypeSN='93' AND DrawingTypeCode='ELSE' and SubEntryId='SEN22P3D' and RecordState='A'"
    cur.execute(sql)
    oldfiles=cur.fetchall()
    i=0
    for (fid,old) in oldfiles:
        new="S83"+old[3:]
        old=old.upper()
        new=new.upper()
        sql="update DesignFile set SysFileName='"+new+"',DrawingTypeSN=83 WHERE Id='"+fid +"'"
        print(sql)
        cur.execute(sql)
        sql="SELECT ID,SysFileName,SubModuleFileName FROM DesignFile WHERE PrjPhaseId='PPS20BKF' AND MajorCode='S' and FileTypeCode='SingleFrame' and SubModuleFileName is not null"
        cur.execute(sql)
        res=cur.fetchall()
        for (fileid,SysFileName,SubModuleFileName) in res:
            SubModules=SubModuleFileName.split(",")
            old=old.lower()
            if old.upper() in SubModuleFileName.upper().split(","):
                lst=[]
                for module in SubModules:
                    if old==module:
                        lst.append(new)
                    else:
                        lst.append(module)
                new_file_names=",".join(lst)
                sql="update DesignFile set SubModuleFileName='"+new_file_names+"' WHERE Id='"+fileid+"'"
                print(sql)
                cur.execute(sql)
    conn.commit()
    conn.close()


if __name__=="__main__":
    main()
