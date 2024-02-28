from io import StringIO
import os
import re
import pdfplumber

class PDF_Operate():
    def readPdf(inputFile):
        text = []
        with pdfplumber.open(inputFile) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text(x_tolerance=2)

                page_text_list = page_text.split("\n")

                text += page_text_list
        return text

    def saveAs(inputFile, outputFile):
        with open(inputFile, 'rb') as fp1:
            b1 = fp1.read()
        with open(outputFile, 'wb') as fp2:
            fp2.write(b1)

# if __name__ == '__main__':
#     dirUrl = "C:\\Users\\chen-fr\\Desktop\\test2"  # 文件夹目录
#     files = os.listdir(dirUrl)  # 得到文件夹下的所有文件名称
#     n = 1
#     invoiceNoStar = 4
#     orderNoStar = 7
#     msg = {}
#     pdfOperate = PDF_Operate
#     for each in files:
#         print(each)
#         fileUrl = '%s\\%s' % (dirUrl, each)
#         if os.path.isfile(fileUrl):
#             with open(fileUrl, 'rb') as my_pdf:
#                 print(n)
#                 fileCon = pdfOperate.readPdf(my_pdf)
#                 print(fileCon)
#                 fileNum = 0
#                 for fileCon[fileNum] in fileCon:
#                     if re.match('.*Invoice No.', fileCon[fileNum]):
#                         msg['Company Name'] = fileCon[fileNum].replace('Invoice No.', '')
#                     elif re.match('%s\d{8}'%invoiceNoStar, fileCon[fileNum]):
#                         print(fileCon[fileNum], 22)
#                         msg['Invoice No'] = fileCon[fileNum]
#                     elif re.match('%s\d{8}'%orderNoStar, fileCon[fileNum]):
#                         print(fileCon[fileNum], 33)
#                         msg['Order No'] = fileCon[fileNum]
#                     elif re.match('\d{2}.\d{3}.\d{2}.\d{4,5}', fileCon[fileNum]):
#                         print(fileCon[fileNum], 44)
#                         msg['Project No'] = fileCon[fileNum]
#                     fileNum += 1
#                 n += 1
#                 outputFlie = msg['Project No'] + '.pdf'
#                 # outputFlie = msg['Invoice No'] + '-' + msg['Company Name'] + '.pdf'
#                 pdfOperate.saveAs(fileUrl,'%s\\test\\%s' % (dirUrl, outputFlie))
#
#         else:
#             print('无')





