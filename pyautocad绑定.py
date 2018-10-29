# -*- coding: utf-8 -*- 运行python
import os,stat
from pyautocad import Autocad, APoint

def run():
    acad =Autocad(create_if_not_exists=True)
    doc=acad.doc
    print(doc.Name)
    xref_purge(doc)
    if bind_xref_1(doc):
        print("bind_xref_1 successful")
    elif bind_xref_2(doc):
        print("bind_xref_2 successful")
    elif bind_xref_3(doc):
        print("bind_xref_3 successful")
    else:
        print("bind xref all failed")

def test():
    acad = Autocad(create_if_not_exists=True)
    acad.prompt("Hello, Autocad from Python\n")
    print(acad.doc.Name)
    #acad.doc.SendCommand('(load "iCapolUpLoadFun.VLX") ')
    
def doc_load(doc):
    doc.SendCommand(r'(arxload "D:\\CapolCAD\\lsp\\iWCapolPurgeIn.arx") ')
    doc.SendCommand(r'(load "D:\\CapolCAD\\lsp\\bindfix-origin.lsp") ')

def find_file(path,find_path):
    (fpath,fname)=os.path.split(path)
    fpath=os.path.join(find_path,fname)
    if os.path.isfile(fpath):
        return fpath
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

def xref_purge(doc):
    doc_load(doc)
    xref_lst=get_all_xref(doc)
    if xref_lst:
        for (name,path) in xref_lst:
            fpath=find_file(path,doc.path)
            fpath=set_write(fpath)
            if fpath:
                if os.path.getsize(fpath)>1000000:
                    doc.SendCommand('(purgefile "'+fpath+'") ')

def get_all_xref(doc):
    lst=[]
    objs=doc.Blocks
    for i in range(objs.Count):
        if objs.item(i).IsXRef:
            lst.append((objs.item(i).Name,objs.item(i).path))     
    return lst

def get_lv1_xref(doc):
    lst=[]
    a_xref=[]
    objs=doc.ModelSpace
    for (name,path) in get_all_xref(doc):
        a_xref.append(name)
    for i in range(objs.Count):
        if objs.item(i).ObjectName=="AcDbBlockReference":
            if objs.item(i).Name in a_xref:
                lst.append((objs.item(i).Name,objs.item(i).path))
    return lst

def bind_xref_1(doc):
    doc_load(doc)
    doc.SendCommand('(bind-fix) ')
    if not get_all_xref(doc):
        return True
    else:
        return False

def bind_xref_2(doc):
    doc_load(doc)
    xref_lst=get_lv1_xref(doc)
    if xref_lst:
        for (name,path) in xref_lst:
            doc.SendCommand('(command "-xref" "b" "'+name+'") ')
        if not get_all_xref(doc):
            return True
        else:
            return False
    else:
        return True

def bind_xref_3(doc):
    doc_load(doc)
    xref_lst=get_lv1_xref(doc)
    if xref_lst:
        find_path=doc.Path
        for (name,path) in xref_lst:
            fpath=find_file(path,find_path)
            fpath=set_write(fpath)
            if fpath:
                doc2=Autocad().app.documents.open(fpath)
                if bind_xref_1(doc2):
                    doc2.close(True)
                elif binf_xref_2(doc2):
                    doc2.close(True)
                else:
                    print("Can not bind xref: "+doc2.name)
                    doc2.close(False)
        if bind_xref_1(doc):
            return True
        elif bind_xref_2(doc):
            return True
        else:
            print("bind_xref_3 failed")
            return False

if __name__=="__main__":
    run()
