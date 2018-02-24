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
            if(".git" in path):
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