#-*- coding: utf-8 -*-
from io import StringIO
from io import open
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager
# from pdfminer.pdfinterp import process_pdf
import os
import re

def readPdf(inputFile):
    # resource manager
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    # device
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    process_pdf(rsrcmgr, device, inputFile)
    device.close()
    content = retstr.getvalue()
    retstr.close()
    # 获取所有行
    lines = str(content).split("\n")
    return lines

def saveAs(inputFile,outputFile):
    with open(inputFile, 'rb') as fp1:
        b1 = fp1.read()
    with open(outputFile, 'wb') as fp2:
        fp2.write(b1)

if __name__ == '__main__':
    # files = [
    #     "C:\\Users\\chen-fr\\Desktop\\DG invoice\\SPOOL_804654.pdf",
    #     "C:\\Users\\chen-fr\\Desktop\\DG invoice\\SPOOL_804665.pdf",
    #     "C:\\Users\\chen-fr\\Desktop\\DG invoice\\SPOOL_804671.pdf",
    #     "C:\\Users\\chen-fr\\Desktop\\DG invoice\\66.441.22.5510.01.pdf",
    #     "C:\\Users\\chen-fr\\Desktop\\DG invoice\\486318964-Sun Lik Shoes & Metal Co., Ltd..pdf",
    #     "C:\\Users\\chen-fr\\Desktop\\DG invoice\\66.441.22.6366.01.pdf",
    #     "C:\\Users\\chen-fr\\Desktop\\DG invoice\\66.441.22.6521.01.pdf",
    #     "C:\\Users\\chen-fr\\Desktop\\DG invoice\\486319620-TGL (HK) LIMITED.pdf",
    #     "C:\\Users\\chen-fr\\Desktop\\DG invoice\\486319792-Sugi International Limited.pdf",
    #     "C:\\Users\\chen-fr\\Desktop\\DG invoice\\486319783-Slip Enterprises Pty Ltd.pdf",
    #     "C:\\Users\\chen-fr\\Desktop\\DG invoice\\486319782-Skys Creative Co. Limited.pdf"
    #
    # ]


    dirUrl = "C:\\Users\\chen-fr\\Desktop\\New folder"  # 文件夹目录
    files = os.listdir(dirUrl)  # 得到文件夹下的所有文件名称
    n = 1
    invoiceNoStar = 4
    orderNoStar = 7
    msg = {}
    for each in files:
        print(each)
        fileUrl = '%s\\%s' % (dirUrl, each)
        if os.path.isfile(fileUrl):
            with open(fileUrl, 'rb') as my_pdf:
                print(n)
                fileCon = readPdf(my_pdf)
                print(fileCon)
                fileNum = 0
                for fileCon[fileNum] in fileCon:
                    if re.match('.*P. R. China', fileCon[fileNum]):
                        if fileCon[fileNum+2] == 'Invoice ':
                            print(fileCon[fileNum], fileCon[fileNum + 4], 11)
                            msg['Company Name'] = fileCon[fileNum + 4]
                        else:
                            print(fileCon[fileNum],fileCon[fileNum+2],11)
                            msg['Company Name'] = fileCon[fileNum + 2]
                    elif re.match('%s\d{8}'%invoiceNoStar, fileCon[fileNum]):
                        print(fileCon[fileNum], 22)
                        msg['Invoice No'] = fileCon[fileNum]
                    elif re.match('%s\d{8}'%orderNoStar, fileCon[fileNum]):
                        print(fileCon[fileNum], 33)
                        msg['Order No'] = fileCon[fileNum]
                    elif re.match('\d{2}.\d{3}.\d{2}.\d{4,5}', fileCon[fileNum]):
                        print(fileCon[fileNum], 44)
                        msg['Project No'] = fileCon[fileNum]
                    fileNum += 1
                n += 1
                # outputFlie = msg['Project No'] + '.pdf'
                outputFlie = msg['Invoice No'] + '-' + msg['Company Name'] + '.pdf'
                saveAs(fileUrl,'%s\\test\\%s' % (dirUrl, outputFlie))

        else:
            print('无')





