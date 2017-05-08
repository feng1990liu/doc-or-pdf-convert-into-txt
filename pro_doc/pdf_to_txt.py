# -*- coding: utf-8 -*-
from subprocess import call
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.layout import *
from pdfminer.converter import PDFPageAggregator
import fnmatch
import os
import sys
import traceback

reload(sys)
sys.setdefaultencoding('gbk')



def Pdf2Txt(path):
    files = os.listdir(path)
    # 该目下创建一个新目录newdir，用来放转化后的txt文本
    New_dir = os.path.abspath(os.path.join(path, 'newdir'))
    if not os.path.exists(New_dir):
        os.mkdir(New_dir)
    #来创建一个pdf文档分析器
    try:
        for f in files:
            if not fnmatch.fnmatch(f, '*.pdf'):
                continue
            print f.decode('gbk')
            try:
                parser = PDFParser(open(path+'\\'+f,'rb'))
                document = PDFDocument(parser)
            except:
                #处理加密的pdf文档，需要安装qpdf  https://sourceforge.net/projects/qpdf/
                call(r'C:\Python27\qpdf-6.0.0\bin\qpdf --password=%s --decrypt %s %s' % ('', path+'\\'+f, path+'\\'+'e'+f), shell=True)
                parser = PDFParser(open(path+'\\'+"e"+f, 'rb'))
                document = PDFDocument(parser)
            if not document.is_extractable:
                raise PDFTextExtractionNotAllowed
            else:
        # 创建一个PDF资源管理器对象来存储共赏资源
                new_txt_name = f[:-4] + '.txt'
                word_to_txt = os.path.join(os.path.join(path, 'newdir'), new_txt_name).decode('gbk')
                f2 = open(word_to_txt,"wb")
                rsrcmgr=PDFResourceManager()
                laparams=LAParams()
        # 创建一个PDF设备对象
        # device=PDFDevice(rsrcmgr)
                device=PDFPageAggregator(rsrcmgr,laparams=laparams)
        # 创建一个PDF解释器对象
                interpreter=PDFPageInterpreter(rsrcmgr,device)
        # 处理每一页
                for page in PDFPage.create_pages(document):
                    interpreter.process_page(page)
            # 接受该页面的LTPage对象
                    layout=device.get_result()
                    for x in layout:
                        if(isinstance(x,LTTextBoxHorizontal)):
                            f2.write(x.get_text().encode('utf-8')+'\n')
                f2.close();
    except:
        traceback.print_exc()

Pdf2Txt(r'D:\class\junmei\wjm\27_317\shanxi')