# -*- coding: utf-8 -*- 运行python
import os,stat,time
from pyautocad import Autocad, APoint
import pywinauto
import pymysql
import subprocess

def autorun(mode=0):
    while 1:
        dlg=dialog()
        dlg.find()
        if dlg.dlg:
            res=run_bind()
            if res.result!='failed_0':
                dlg.dlg["确定Button"].click()
            else:
                log_result(res.name,res.path,res.result)
                print("try again:")
                res=run_bind()#再试一下
                if res.result!='failed_0':
                    dlg.dlg["确定Button"].click()
                else:
                    dlg.dlg["取消Button"].click()
                    print("取消Button")
            log_result(res.name,res.path,res.result)
        time.sleep(1)
        print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime()))
        
def run_bind():
    bind=binding()
    bind.get_doc()
    print(bind.name)
    bind.close_other()
    bind.xref_purge()
    bind.bind_xref_1()
    if bind.has_xref():
        bind.remove_unload_xref()
        bind.bind_xref_2()
    else:
        bind.result='success_1'
        return bind
    if bind.has_xref():
        bind.bind_xref_3()
    else:
        bind.result='success_2'
        return bind
    if bind.has_xref():
        bind.result='failed_0'
        return bind
    else:
        bind.result='success_3'
        return bind

def print_func(func):
    def inner(*args,**kwargs):
        print(func.__name__+':')
        print(args,kwargs)
        return func(*args,**kwargs)
    return inner

class dialog():
    def __init__(self):
        self.app=pywinauto.application.Application()
        self.dlg=0

    def find(self):
        self.app.connect(path="D:\\AcsModule\\Client\\AcsTools.exe")
        try:
            self.dlg =self.app[r"提示"]
            self.dlg[r"存在无法绑定的外参，手动绑定完成后点[确定],放弃绑定点[取消]！"].print_control_identifiers()
        except pywinauto.findbestmatch.MatchError:
            self.dlg=0
        return self.dlg


class binding():
    def __init__(self):
        self.acad =Autocad(create_if_not_exists=True)
        self.acad.Application.LoadARX("D:\\CapolCAD\\lsp\\iWCapolPurgeIn.arx")
        self.result='failed_0'

    def get_doc(self):
        self.doc=self.acad.doc
        self.path,self.name=self.doc.Path,self.doc.Name
        return self.doc

    def has_xref(self):
        self.get_all_xref()
        if self.a_xref_name:
            return True
        else:
            return False

    def get_all_xref(self):#获取所有外参
        lst=[]
        self.a_xref_name=[]
        self.a_xref_path=[]
        self.a_xref_name_path=[]
        self.Blocks=self.doc.Blocks
        for i in range(self.Blocks.Count):
            obj=self.Blocks.item(i)
            if obj.IsXRef:
                lst.append(obj) 
                xname,xpath=obj.Name,obj.Path
                self.a_xref_name.append(xname)
                self.a_xref_path.append(xpath)
                self.a_xref_name_path.append((xname,xpath))
        return lst

    @print_func
    def get_lv1_xref(self):#获取一级外参
        lst=[]
        self.lv1_xref_name=[]
        self.lv1_xref_path=[]
        self.lv1_xref_name_path=[]
        self.get_all_xref()
        self.ModelSpace=self.doc.ModelSpace
        for i in range(self.ModelSpace.Count):
            obj=self.ModelSpace.item(i)
            if obj.ObjectName=="AcDbBlockReference":
                if obj.Name in self.a_xref_name:
                    lst.append(obj)
                    xname,xpath=obj.Name,obj.Path
                    self.lv1_xref_name.append(xname)
                    self.lv1_xref_path.append(xpath)
                    self.lv1_xref_name_path.append((xname,xpath))
        return lst

    def load(self):
        self.doc.SendCommand(r'(load "D:\\CapolCAD\\lsp\\bindfix-origin.lsp") ')

    @print_func
    def close_other(self):
        self.get_doc()
        lst=[]
        docs=Autocad(create_if_not_exists=True).app.documents
        for i in range(docs.count):
            if docs[i].Name != self.name:
                lst.append(docs[i].Name)
        for doc in lst:
            docs[doc].close()

    @print_func
    def xref_purge(self):
        self.load()
        self.get_all_xref()
        if self.a_xref_path:
            for path in self.a_xref_path:
                if path and path!="D:\\AcsModule\\Client\\EmptyDwg.dwg": 
                    fpath=find_file(path,self.path)
                    if fpath:
                        if os.path.getsize(fpath)>1000000:
                            print("purge file: " +fpath)
                            fpath=set_write(fpath)
                            self.purgefile(fpath)
    
    @print_func
    def purgefile(self,fpath):
        self.doc.SendCommand('(purgefile "'+fpath+'") ')

    def count0_lst(self):#获取未参照
        lst=[]
        objs=self.doc.Blocks
        for i in range(objs.Count):
            if objs.item(i).IsXRef:
                if objs.item(i).Count==0:
                    lst.append(objs.item(i))
        return lst

    @print_func
    def remove_count0(self):#拆离未参照
        lst=self.count0_lst()
        if lst:    
            for item in lst:
                self.doc.SendCommand('(command "-xref" "d" "'+item.Name+'") ')
            self.doc.SendCommand('(command "-xref" "r" "*") ')
            lst=self.count0_lst()
            return lst
        else:
            return False
    '''
    def remove_unload_xref(self):
        for i in range(5):
            if not self.remove_count0():
                break 
        if self.count0_lst():
            fpath=os.path.join(self.path,self.name)
            print(fpath)
            if fpath:
                self.doc.close(True)
                new_doc=self.open_file(fpath)
                if new_doc:
                    return new_doc
                else:
                    return False
            else:
                return False
        else:
            return self
        '''

    @print_func
    def remove_unload_xref(self):
        for i in range(5):
            if not self.remove_count0():
                return True
        return False
        
    @print_func
    def bind_xref_1(self):#快速绑定
        self.load()
        self.doc.SendCommand('(bind-fix) ')
        self.get_all_xref()

    @print_func
    def bind_xref_2(self):#逐个绑定
        self.load()
        self.get_lv1_xref()
        if self.lv1_xref_name:
            self.doc.SendCommand('(command "-xref" "r" "*") ')
            for name in self.lv1_xref_name:
                self.doc.SendCommand('(command "-xref" "b" "'+name+'") ')
            self.get_all_xref()

    @print_func
    def bind_xref_3(self):#一级深度绑定
        self.load()
        self.get_lv1_xref()
        if self.has_xref():
            for path in self.lv1_xref_path:
                fpath=find_file(path,self.path)
                fpath=set_write(fpath)
                if fpath:
                    doc2=open_file(fpath)
                    #doc2.doc=Autocad().app.documents.open(fpath)
                    #doc2=binding()
                    #doc2.get_doc()
                    if doc2:
                        doc2.bind_xref_1()
                        if doc2.has_xref():
                            doc2.bind_xref_2()
                        if doc2.has_xref():
                            print("Can not bind xref: "+doc2.name)
                            doc2.doc.close(False)
                        else:
                            doc2.doc.close(True)
            self.doc.SendCommand('(command "-xref" "r" "*") ')       
            self.bind_xref_1()
            if self.has_xref():
                self.bind_xref_2()

def find_file(path,find_path):
    if path and find_path:
        (fpath,fname)=os.path.split(path)
        fpath=os.path.join(find_path,fname)
        if os.path.isfile(fpath):
            return fpath
        else:
            return ""
    else:
        return ""

def set_write(path):
    try:
        with open(path,"r+") as fr:
            return path
    except PermissionError:
        try:
            os.chmod(path, stat.S_IWRITE)
            return path
        except:
            print("Can not set_write:"+path)
            return False
'''
def open_file(path):
    new_doc=binding()
    new_doc.doc=new_doc.acad.Application.documents.open(path)
    return new_doc
'''
@print_func
def open_file(path):
    new_doc=binding()
    try:
        new_doc.doc=new_doc.acad.Application.documents.open(path)
        return new_doc
    except:
        with open('D:\\CapolCAD\\lsp\\open_file.bat','w') as f:
            f.write('"'+path+'"')
        run_bat=os.system('D:\\CapolCAD\\lsp\\open_file.bat')
        #new_doc.get_doc()
        #new_doc.doc.SendCommand('CAPOL_OPENDWG ')
        '''
        filedia=new_doc.doc.GetVariable("filedia")
        new_doc.doc.SetVariable("filedia",0)
        new_doc.doc.SendCommand('_open %s ' % path)
        #new_doc.doc.SendCommand(path + " ")
        new_doc.doc.SetVariable("filedia",filedia)
        '''
        for i in range(60):
            time.sleep(1)
            try:
                new_doc.get_doc()
                print(os.path.join(new_doc.path,new_doc.name))
                if path.upper() == os.path.join(new_doc.path,new_doc.name).upper():
                    return new_doc
            except:
                pass
    print("open file time out!")
    return 0

@print_func
def log_result(file_name,file_path,result):
    print("#".join(['file name',file_name,'file path',file_path,'result',result]))
    conn=pymysql.connect(host="10.1.42.24",user="steven",passwd="steven!@#456",db="project_test")
    cur=conn.cursor()
    cur.execute(u"INSERT INTO BindHelperRecord(file_name,file_path,result) VALUES(%s,%s,%s)",[file_name,file_path,result])
    conn.commit()
    cur.close()
    conn.close()

if __name__=="__main__":
    autorun()
    #test()
    #run_bind()
    #docs=Autocad(create_if_not_exists=True).app.documents
    #for i in range(docs.count):
        #print(docs[i].Name)
    #bind=binding()
    #bind.close_other()


