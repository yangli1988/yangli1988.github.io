import os
import docx
from win32com import client as wc
import HTMLParser
import shutil
import chardet

def walkall_files(dirpath):
    fps = []
    for (root,subdirs,files) in os.walk(dirpath):
        for fn in files:
            path = os.path.join(root,fn)
            if(".git\\" in path):
                pass
            else:
                fps.append(path)
    return(fps)


def new_paths(fps):
    nfps = []
    for i in range(0,fps.__len__()):
        np = fps[i].replace("\\Lawyer\\","\\Lawyer2\\")
        nfps.append(np)
    return(nfps)


# def read_docx(file_name):
    # doc = docx.Document(file_name)
    # content = '\n'.join([para.text for para in doc.paragraphs])
    # return(content)

def doc2docx(src_path,saveas_path,word_app ="Word.Application"):
    word = wc.Dispatch(word_app)
    #因为如果遇到错误文档会卡住，此处需要优化加入error处理 OpenAndRepair
    # word = DispatchEx('Word.Application') #启动独立的进程
    word.Visible = 0  # 后台运行,不显示
    word.DisplayAlerts = 0  # 不警告
    #doc = word.Documents.Open(FileName=path, Encoding='gbk')
    try:
        doc = word.Documents.Open(src_path)
        #使用参数16表示将doc转换成docx
        doc.SaveAs(saveas_path,16)
        doc.Close()
    except:
        return(False)
    else:
        return(True)
    word.Quit()

def doc2x(fps):
    for i in range(0,fps.__len__()):
        path = fps[i]
        suffix = os.path.splitext(path)
        if(suffix[1] == ".doc"):
            print(fps[i])
            new_path = fps[i]+"x"
            succ = doc2docx(fps[i],new_path)
            if(succ):
                os.remove(fps[i])
                fps[i] = new_path
            else:
                error_path = fps[i]+".error"
                os.rename(fps[i],error_path)
                fps[i] = error_path
        else:
            pass
    return(fps)

fps = walkall_files(os.getcwd() + "\\Lawyer")
fps = doc2x(fps)
nfps = new_paths(fps)



def docx2html(src_path,saveas_path,word_app ="Word.Application"):
    word = wc.Dispatch(word_app)
    #因为如果遇到错误文档会卡住，此处需要优化加入error处理 OpenAndRepair
    # word = DispatchEx('Word.Application') #启动独立的进程
    word.Visible = 0  # 后台运行,不显示
    word.DisplayAlerts = 0  # 不警告
    #doc = word.Documents.Open(FileName=path, Encoding='gbk')
    try:
        doc = word.Documents.Open(src_path)
        #8 是html
        # .html文件 连同 files文件夹会自动保存
        doc.SaveAs(saveas_path,8)
        doc.Close()
    except:
        word.Quit()
        return(False)
    else:
        word.Quit()
        return(True)
    


def x2html(fps,nfps):
    for i in range(0,fps.__len__()):
        dirname = os.path.dirname(nfps[i])
        try:
            os.makedirs(dirname)
        except:
            pass
        else:
            pass
        path = fps[i]
        npath = nfps[i]
        suffix = os.path.splitext(path)
        nsuffix = os.path.splitext(npath)
        if(suffix[1] == ".docx"):
            print(fps[i])
            new_path = nsuffix[0]+".html"
            succ = docx2html(fps[i],new_path)
            if(succ):
                nfps[i] = new_path
            else:
                error_path = fps[i]+".error"
                os.rename(fps[i],error_path)
                fps[i] = error_path
                shutil.copy(fps[i],nfps[i])
        else:
            shutil.copy(fps[i],nfps[i])
    return(nfps)

nfps = x2html(fps,nfps)


###

import os
import docx
from win32com import client as wc
import HTMLParser
import shutil


def walkall_files(dirpath):
    fps = []
    for (root,subdirs,files) in os.walk(dirpath):
        for fn in files:
            path = os.path.join(root,fn)
            if(".git\\" in path):
                pass
            elif(".files" in path):
                pass
            else:
                fps.append(path)
    return(fps)

fps = walkall_files(os.getcwd() + "\\Lawyer")

def prepend_to_file(prepend,**kwargs):
    prepend=bytes(prepend)
    fd = open(kwargs['fn'],"rb+")
    rslt = fd.read()
    fd.close()
    os.remove(kwargs['fn'])
    fd = open(kwargs['fn'],"wb+")
    fd.write(prepend+rslt)
    fd.close()


def insert_doctype(fps):
    doctype = b"<!DOCTYPE html>\r\n"
    length = fps.__len__()
    for i in range(0,length):
        path = fps[i]
        if(".html" in path):
            prepend_to_file(doctype,fn=path)
        else:
            pass
    return(fps)

fps = insert_doctype(fps)

def read_file_content(**kwargs):
    fd = open(kwargs['fn'],kwargs['op'])
    rslt = fd.read()
    fd.close()
    return(rslt)
    
    
    
import chardet

def convert_code(to_codec="UTF8",**kwargs):
    fd = open(kwargs['fn'],"rb+")
    rslt = fd.read()
    fd.close()
    from_codec = chardet.detect(rslt)['encoding']
    if(from_codec == "utf-8"):
        pass
    else:
        rslt = rslt.decode(from_codec,'ignore').encode(to_codec)
        os.remove(kwargs['fn'])
        fd = open(kwargs['fn'],"wb+")
        fd.write(rslt)
        fd.close()

def detect_code(fn):
    fd = open(fn,"rb+")
    rslt = fd.read()
    fd.close()
    from_codec = chardet.detect(rslt)['encoding']
    print(from_codec)
    return(rslt)



def convert_all(fps):
    '''to solve chinese display bug'''
    failed =[]
    length = fps.__len__()
    for i in range(0,length):
        path = fps[i]
        if(".html" in path):
            try:
                convert_code(to_codec="UTF8",fn=path)
            except:
                failed.append(path)
                print(path)
            else:
                pass
        else:
            pass
    return(failed)

failed = convert_all(fps)


#####################


def walkall_dirs(dirpath):
    dirs = []
    for (root,subdirs,files) in os.walk(dirpath):
        for subdir in subdirs:
            path = os.path.join(root,subdir)
            if(".git\\" in path):
                pass
            elif(".files" in path):
                pass
            else:
                dirs.append(path)
    return(dirs)

dirs = walkall_dirs(os.getcwd() + "\\Lawyer")


def get_urls(fps):
    urls = []
    length = fps.__len__()
    for i in range(0,length):
        url = fps[i].replace("\\","/").replace("D:/LiYang/","https://")
        urls.append(url)
    return(urls)

urls = get_urls(fps)


def category(fps):
    categ = {}
    length = fps.__len__()
    for i in range(0,length):
        dir = os.path.dirname(fps[i])
        if(dir in categ):
            categ[dir].append(fps[i])
        else:
            categ[dir] = [fps[i]]
    return(categ)

categ = category(fps)


#对与每个dir 要生成一个index.html






def creat_indexes_html(path,dirname,url):
    
    rslt= "<html>\r\n<head><\head><body>\r\n"
    rslt = rslt + '<a href="' + url +'">'
    rslt = rslt + os.path.splitext(fn)[0] + "<\a>\r\n"
    rslt = rslt + "</body>\r\n+<\html>"
    
    

