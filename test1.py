# -*- coding: utf-8 -*-
import pymssql,cx_Oracle
import os



def run():
    lst=[
    ('深圳','10.1.246.1'),
    ('广州','10.2.1.114'),
    ('长沙','10.3.1.3'),
    ('上海','10.6.1.240')
    ]
    #oracle='steven/qfc23834358@10.1.42.66:1521/xe'
    oracle='ccd/hygj1234@10.1.1.213:1521/orcl'
    #drop_and_create(oracle)
    with cx_Oracle.connect(oracle) as conn:
        cur=conn.cursor()
        cur.execute(u"""select 工程编号,工程名称,工程阶段,专业,子项号,子项名称,图号,图纸名称,图纸分项 as 图纸类型,图纸子类,设计阶段,校审阶段,意见等级,
case when 意见状态=0 then  '未处理'
when 意见状态=1 then  '部分修改'
when 意见状态=2 then  '已修改'
when 意见状态=3 then  '有争议'
when 意见状态=4 then  '后续修改'
when 意见状态=5 then  '无效意见'
when 意见状态=6 then  '复核通过'
when 意见状态=7 then  '复核驳回'
else '' end as 意见状态,
校审人,校审意见,校审时间,意见责任人,反馈意见,反馈时间,复审意见,复审时间,复审次数
from Capol_ProjectOpinions_2018""")
        print(cur.fetchall())



if __name__=="__main__":
    run()
