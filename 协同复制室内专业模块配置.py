# -*- coding: utf-8 -*- 运行python
import pymssql
from openpyxl import load_workbook
import os

def copy_md_setting(cur,src_ppid,tar_ppid,MajorCode):
    cur.execute(u"DELETE FROM [dbo].[MD_DrawingType] WHERE PrjPhaseId=%s and MajorCode=%s",(tar_ppid,MajorCode))
    cur.execute(u"select ID from md_drawingtype where PRJPHASEID=%s AND MAJORCODE=%s",(src_ppid,MajorCode))
    for (src_md_id,) in cur.fetchall():
        print(src_md_id,src_ppid,MajorCode,tar_ppid)
        cur.execute("""INSERT INTO [dbo].[MD_DrawingType]([Id],[PrjPhaseId],[DrawingTypeSN],[DrawingTypeCode],[DrawingTypeText],[DrawingTypeIndex],[MajorCode],
        [MajorName],[FileTypeCode],[FileTypeText],[DesignClassCode],[DesignClassText],[IsIncludeFloorName],[TakeoutMajorCodes],
        [TakeoutMajorNames],[IsUseRootSN],[IsOutMajorMuilteModule],[IsTakeout],[SyncTime],[IsTakeoutFrame],
        [IsTakeoutFrameModule],[PrjShareType],[IsShowPrjShareCatalog],[IsShowSingleModule],[GroupSN],[GroupName])
        SELECT [Id] , %s as [PrjPhaseId],[DrawingTypeSN],[DrawingTypeCode],[DrawingTypeText],[DrawingTypeIndex],[MajorCode],
        [MajorName],[FileTypeCode],[FileTypeText],[DesignClassCode],[DesignClassText],[IsIncludeFloorName],[TakeoutMajorCodes],
        [TakeoutMajorNames],[IsUseRootSN],[IsOutMajorMuilteModule],[IsTakeout],[SyncTime],[IsTakeoutFrame],
        [IsTakeoutFrameModule],[PrjShareType],[IsShowPrjShareCatalog],[IsShowSingleModule],[GroupSN],[GroupName]FROM [dbo].[MD_DrawingType] where [Id]=%s and [PrjPhaseId]=%s""",(tar_ppid,src_md_id,src_ppid))

def run():
    with pymssql.connect(host="10.1.246.1", user="sa", password="sasasa",database="CAPOL_Project") as conn:  #深圳新协同
        cur=conn.cursor()
        src_ppid='PPS20BJ8'
        tar_ppid='PPS20BJX'
        MajorCode='I'
        copy_md_setting(cur,src_ppid,tar_ppid,MajorCode)
       
       # conn.commit()


if __name__=="__main__":
    run()