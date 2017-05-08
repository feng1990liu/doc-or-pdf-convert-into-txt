#encoding=utf-8
# import docx
import re
import codecs
import csv
import os.path
import sys
import time
reload(sys)
sys.setdefaultencoding('utf-8')
##定义读当前目录
# 遍历指定目录，显示目录下的所有文件名:docx


def getFileList(path1, errorList):
    listFile = os.listdir(path1)
    infoList = []
    print len(listFile)
    for i in listFile:
        fileName = i.decode('gbk')
        try:
            txt = readTxt(path1, fileName)
        except:
            errorList.append(path1+fileName)
            continue
        curFileInfo = getEntryPertxt(txt,fileName)  ##记录每个文档各个项目对应的值
        infoList.append(curFileInfo)
    return infoList

def getEntryPertxt(txt,fileName):
    curFileInfo = []  ##记录每个文档各个项目对应的值
    fileNameTrim = fileName.split('.')[0]  # re.search('.*?\.',fileName)  #公司名称'
    curFileInfo.append(fileNameTrim)

    IntrPat1 = ur'(前.*?言)'
    # 第1章
    qualityPat = ur'(企业质量理念)'
    # 第2章   企业质量管理
    quaManaPat = ur'(企业质量管理)'
    instiPat = ur'(质量管理机构)'
    tixiPat = ur'(质量管理体系)'
    riskPat = ur'(质量安全风险管理)'

    # 第3章  质量诚信管理:质量承诺	运作管理	营销管理
    intePat = ur'(质量诚信管理)'
    promisePat = ur'(质量承诺)'
    operaPat = ur'(运作管理)'
    marketPat = ur'(营销管理)'

    ##第4章：4.质量管理基础 ：标准管理	计量管理	认证管理	检验检测管理
    basisPat = ur'(质量管理基础)'
    critePat = ur'(标准管理)'
    countPat = ur'(计量管理)'
    authPat = ur'(认证管理)'
    checkPat = ur'(检验检测管理)'

    ##5.产品质量责任	: 产品质量水平	产品售后责任	企业社会责任	质量信用记录
    dutyPat = ur'产品质量责任'
    levelPat = ur'产品质量水平'
    afSalePat = ur'产品售后责任'
    esrPat = ur'企业社会责任'  # Enterprise social responsibility
    recordPat = ur'质量信用记录'

    ##结束语
    endPat1 = ur'结.*?语'
    endPat2 = ur'结.*?束.*?语'
    patList = [IntrPat1,qualityPat,quaManaPat,instiPat, tixiPat, riskPat,intePat,promisePat, operaPat, marketPat,
               basisPat,critePat, countPat, authPat, checkPat,dutyPat,levelPat, afSalePat, esrPat, recordPat]

    #  '前言','1.企业质量理念','质量管理机构','质量管理体系','质量安全风险管理','质量承诺','运作管理',
    # '营销管理','标准管理','计量管理','认证管理','检验检测管理','产品质量水平','产品售后责任','企业社会责任','质量信用记录'
    # 针对每一章作处理
    for j in patList:
        patj = re.compile(j)
        s1 = re.search(patj, txt)  #而re.search匹配整个字符串，直到找到一个匹配。
        if s1:
            # s1 = s1.group()
            curFileInfo.append('√')
        else:
            curFileInfo.append(0)

    pend1 = re.compile(endPat1)
    pe1 = re.search(pend1,txt)
    pend2 = re.compile(endPat2)
    pe2 = re.search(pend2, txt)
    if pe1 or pe2:
        curFileInfo.append('√')
    else:
        curFileInfo.append(0)
    count = -1
    if 0 in curFileInfo:
        count = curFileInfo.count(0)
    else:
        count = u'合格'
    curFileInfo.append(count)
    return curFileInfo

##定义读取txt文档
def readTxt(filepath,docName):
    path = os.path.join(filepath, docName)
    print docName
    try:
        f = codecs.open(path, "rb",'gbk')
        fullText = f.readlines()
    except:
        f = codecs.open(path, "rb", 'utf-8')
        fullText = f.readlines()

    str=''
    for j in fullText:
        str = str + j
    # print str[:100]
    return str



def writeStdXls(infoList):
    path = r"D:\class\junmei\wjm\27_317\shanxi\newdir\stdResult2.csv"
    # f = open(path,"wb")
    # data =''
    # f.write(codecs.BOM_UTF8)
    csvFile = file(path, 'wb')
    csvFile.write(codecs.BOM_UTF8)
    writer = csv.writer(csvFile,dialect='excel')
    headers = ['公司名称','前言','2.企业质量管理','企业质量理念','质量管理机构','质量管理体系','质量安全风险管理','3.质量诚信管理','质量承诺','运作管理',
               '营销管理','4.质量管理基础	','标准管理','计量管理','认证管理','检验检测管理','5.产品质量责任','产品质量水平','产品售后责任','企业社会责任','质量信用记录','结束语','Flag']

    writer.writerow(headers)
    for i in infoList:
        writer.writerow(i)
    csvFile.close()
