#encoding=utf-8
import pro_doc.docProcess as  dp

if __name__ == '__main__':
    errorList = []
    # filePath = r'E:\\918\4-12\\wjm\\27_317\\'
    path1 = r'D:\class\junmei\wjm\27_317\newdir'
    info = dp.getFileList(path1,errorList)
    dp.writeStdXls(info)
    for i in errorList:
        print i