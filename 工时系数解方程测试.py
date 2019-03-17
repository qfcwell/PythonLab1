# -*- coding: utf-8 -*-
import pymssql
import numpy as np
from numpy.linalg import solve
'''
a=np.mat([[2,3],[1,3]])#系数矩阵
b=np.mat([0,0]).T    #常数项列矩阵
x=solve(a,b)        #方程组的解
print(x)
'''

def sv(x1,x2,y1,y2):
    a = np.array([[x1, x2], [1, 1]])
    b = np.array([y1, y2])
    return solve(a, b)


with pymssql.connect(host='10.1.42.103', user="sa", password="qfc23834358Q",database="test2") as conn:
    cur=conn.cursor()
    #cur.execute("select [工程编号],[小计],[施工服务],[填报工时],[总预算工时] from wh_table2 where [工程编号] in ('GC150235','GC150082','GC160217','GC150067','GC170374','GC170117') order by [工程编号] asc")
    #cur.execute("select [工程编号],[小计],[施工服务],[填报工时],[总预算工时] from wh_table2 where [工程编号] in ('GC150235','GC160217','GC150067','GC170374') order by [工程编号] asc")
    cur.execute("select [工程编号],[小计],[施工服务],[总预算工时]*[总工时系数],[总预算工时] from wh_table2 where [阶段组合] IN ('初设+施工图+施工服务','方案+施工图+施工服务','方案+初设+施工图+施工服务','施工图+施工服务','方案配合+初设+施工图+施工服务','方案配合+初设配合+施工图+施工服务') and [施工服务]>0 and [小计]>0.2 and [工程编号] in (select fd_project_no from [10.1.1.117].cip.dbo.wt_project_info where doc_create_time>'2015-01-01' and doc_company_id in (select fd_id from [10.1.1.117].cip.dbo.sys_org_element where fd_ldap_dn='OU=深圳公司,OU=华阳设计,OU=华阳国际,DC=capol,DC=cn')) order by [工程编号] asc")
    res=cur.fetchall()
    a_y=b_y=0
    z_t=z_y=0
    a_t=b_t=0
    for (gcno,x1,x2,y1,y2) in res:
        x=sv(x1,x2,y1,y2)
        print(gcno)
        print(x)
        a_y+=x[0]
        b_y+=x[1]
        z_t+=y1
        z_y+=y2
        a_t+=x[0]*x1
        b_t+=x[1]*x2
    r={'设计预算':a_y,'设计填报':a_t,'施工预算':b_y,'施工填报':b_t,'总预算':z_y,'总填报':z_t}
    print(r)
    
    print(r['设计预算']+r['施工预算'])
    print(r['总预算'])
    print(r['设计填报']+r['施工填报'])
    print(r['总填报'])
    print(r['设计填报']/r['设计预算'])
    print(r['施工填报']/r['施工预算'])


