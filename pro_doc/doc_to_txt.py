# -*- coding: cp936 -*-

# 注意Windows下路径表示
from win32com import client as wc
import os
import fnmatch
import sys
reload(sys)
sys.setdefaultencoding('gbk')
all_FileNum = 0
debug = 0
errorList = []

def Translate(path):
    '''''
    将一个目录下所有doc和docx文件转成txt
    该目录下创建一个新目录newdir
    新目录下fileNames.txt创建一个文本存入所有的word文件名
    本版本具有一定的容错性，即允许对同一文件夹多次操作而不发生冲突
    '''
    global debug, all_FileNum
    if debug:
        print path
        # 该目录下所有文件的名字
    files = os.listdir(path)
    print files
    # 该目下创建一个新目录newdir，用来放转化后的txt文本
    New_dir = os.path.abspath(os.path.join(path, 'newdir'))
    if not os.path.exists(New_dir):
        os.mkdir(New_dir)
    if debug:
        print New_dir
        # 创建一个文本存入所有的word文件名
    fileNameSet = os.path.abspath(os.path.join(New_dir, 'fileNames.txt'))
    o = open(fileNameSet, "w")
    wordapp = wc.Dispatch('Word.Application')
    try:
        for filename in files:
            if debug:
                print filename
                # 如果不是word文件：继续
            if not (fnmatch.fnmatch(filename, '*.doc') or fnmatch.fnmatch(filename, '*.docx') or fnmatch.fnmatch(filename, '*.wps') or fnmatch.fnmatch(filename, '*.dot')):
                continue
                # 如果是word临时文件：继续
            if fnmatch.fnmatch(filename, '~$*'):
                continue
            if debug:
                print filename
            docpath = os.path.abspath(os.path.join(path, filename))

            # 得到一个新的文件名,把原文件名的后缀改成txt
            new_txt_name = ''
            if fnmatch.fnmatch(filename, '*.doc') or fnmatch.fnmatch(filename, '*.wps') or fnmatch.fnmatch(filename, '*.dot'):
                new_txt_name = filename[:-4] + '.txt'
            else:
                new_txt_name = filename[:-5] + '.txt'
            if debug:
                print new_txt_name
            word_to_txt = os.path.join(os.path.join(path, 'newdir'), new_txt_name).decode('gbk')
            print word_to_txt

            # 为了让python可以在后续操作中r方式读取txt和不产生乱码，参数为4
            try:
                doc = wordapp.Documents.Open(docpath)
                doc.SaveAs(word_to_txt,4)
                doc.Close()
            except:
                errorList.append(word_to_txt)
    finally:
        wordapp.Quit()



if __name__ == '__main__':
    Translate(r'D:\class\junmei\wjm\27_317')
    print 'The Total Files Numbers = ', all_FileNum
    for i in errorList:
        print i.decode('gbk')