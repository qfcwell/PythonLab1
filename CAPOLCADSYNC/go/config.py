# -*- coding: utf-8 -*-
import os

#os.environ['ORACLE_HOME'] = 'D:\\PLSQL\\instantclient_11_2'
#os.environ['TNS_ADMIN'] = 'D:\\PLSQL\\instantclient_11_2\\network\\admin'
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'

acs_server_lst={
    '深圳':'szacsdb.capol.cn',
    '广州':'gzacsdb.capol.cn',
    '长沙':'csacsdb.capol.cn',
    '上海':'shacsdb.capol.cn',
    '海南':'10.14.2.10'
    }
conn_str='ccd/hygj1234@10.1.1.213:1521/orcl'#capol

if __name__=="__main__":
    pass
