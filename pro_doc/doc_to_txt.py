# -*- coding: cp936 -*-

# ע��Windows��·����ʾ
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
    ��һ��Ŀ¼������doc��docx�ļ�ת��txt
    ��Ŀ¼�´���һ����Ŀ¼newdir
    ��Ŀ¼��fileNames.txt����һ���ı��������е�word�ļ���
    ���汾����һ�����ݴ��ԣ��������ͬһ�ļ��ж�β�������������ͻ
    '''
    global debug, all_FileNum
    if debug:
        print path
        # ��Ŀ¼�������ļ�������
    files = os.listdir(path)
    print files
    # ��Ŀ�´���һ����Ŀ¼newdir��������ת�����txt�ı�
    New_dir = os.path.abspath(os.path.join(path, 'newdir'))
    if not os.path.exists(New_dir):
        os.mkdir(New_dir)
    if debug:
        print New_dir
        # ����һ���ı��������е�word�ļ���
    fileNameSet = os.path.abspath(os.path.join(New_dir, 'fileNames.txt'))
    o = open(fileNameSet, "w")
    wordapp = wc.Dispatch('Word.Application')
    try:
        for filename in files:
            if debug:
                print filename
                # �������word�ļ�������
            if not (fnmatch.fnmatch(filename, '*.doc') or fnmatch.fnmatch(filename, '*.docx') or fnmatch.fnmatch(filename, '*.wps') or fnmatch.fnmatch(filename, '*.dot')):
                continue
                # �����word��ʱ�ļ�������
            if fnmatch.fnmatch(filename, '~$*'):
                continue
            if debug:
                print filename
            docpath = os.path.abspath(os.path.join(path, filename))

            # �õ�һ���µ��ļ���,��ԭ�ļ����ĺ�׺�ĳ�txt
            new_txt_name = ''
            if fnmatch.fnmatch(filename, '*.doc') or fnmatch.fnmatch(filename, '*.wps') or fnmatch.fnmatch(filename, '*.dot'):
                new_txt_name = filename[:-4] + '.txt'
            else:
                new_txt_name = filename[:-5] + '.txt'
            if debug:
                print new_txt_name
            word_to_txt = os.path.join(os.path.join(path, 'newdir'), new_txt_name).decode('gbk')
            print word_to_txt

            # Ϊ����python�����ں���������r��ʽ��ȡtxt�Ͳ��������룬����Ϊ4
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