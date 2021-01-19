
''' Imports '''
try:
    import re2 as re
except ImportError:
    import re
import json
from PyPDF2 import PdfFileWriter, PdfFileReader
from datetime import date
import glob 
import os
import shutil
from prettytable import PrettyTable
from pathlib import Path
import schedule
import time

def main():

    today = date.today()
    extraction_info = PrettyTable()
    extraction_info.field_names = ["pdf file name", "num of pdfs extracted"]

    ogpath = 'C:\\Users\\0235124\\Desktop\\signodeProjects\\pdfParse\\pyPDF2'
    path = ogpath + '\\OneDrive - University of Waterloo\\' + str(today)
    emailArchivePath = ogpath + '\\email_archive'
    pdf_files = glob.glob(os.path.join(ogpath,'*.pdf'))
    

    def pickTicketRawInfo(info):
        pattern = re.compile(r'Order(.*)FOB')
        patternMatches = pattern.finditer(info)
        pickTicketData = [match for match in patternMatches] # text data
        return pickTicketData

    def orderNo (info):
        pattern = re.compile(r'(\d{4,7})[-](\d{2})')
        patternMatches = pattern.finditer(info)
        pickTicketData = [match for match in patternMatches]
        return pickTicketData[0]

    def createDir(path):
        try:
            os.mkdir(path)
            print('folder created', os.path.basename(path))
        except OSError:
            print ("Creation of the directory %s failed" % path)
        # else:
            # print ("Successfully created the directory %s " % path)

    def enterDir(path):
        try: 
            os.chdir(path)
            # print("Directory changed")
        except OSError:
            try: 
                # createDir(path)
                pass
            except:
                print("Can't change the Current Working Directory (Check if the dir was created)")        
                print("Current Working Directory " , os.getcwd())

    def returnToOgPath() : os.chdir(ogpath)

    def prettyInfo(pdf, pages):
        extraction_info.add_row([os.path.basename(pdf), pages])

    def save_Rename_Pdf(pages, file):   
        for page in range(pages):
            pageInfo = inputPdf.getPage(page)
            rawData = pageInfo.extractText() # extract initial data
            rawData = re.sub(r'\r\n', ' ', rawData) # substitute \n with ' '
            rawData = pickTicketRawInfo(rawData)
            orders.append(orderNo(rawData[0].group(0)).group(0))
            # masterData.append(rawData)
            # temp.append(masterData[i][0].group(0))
            # tempdata.append(orderNo(temp[i]).group(0))
            output = PdfFileWriter()
            output.addPage(pageInfo)
            with open("%s.pdf" %(orders[page]), "wb") as outputStream: # create a orderNo list to extract each # (orderNo(temp[i]).group(0))
                output.write(outputStream)
        
        prettyInfo(file, pages)
        # print('path of pdfFile extracted ########', file, '########')

    def moveEmailArchive(src, dst):
        os.replace(src, dst)
        # Path(src).rename(dst)
        

    for pdf_file in pdf_files:
        openPdf = open(pdf_file, "rb")
        inputPdf = PdfFileReader(openPdf) 
        noPages = inputPdf.numPages

        masterData = []
        orders = []

        archivedPdfPath = emailArchivePath + '\\extracted_' + os.path.basename(pdf_file) 

        pathCheck = os.path.exists(path)
        emailPathCheck = os.path.exists(emailArchivePath)
        if not (pathCheck and emailPathCheck):
            createDir(path)
            createDir(emailArchivePath)

        enterDir(path)
        save_Rename_Pdf(noPages, pdf_file)
        returnToOgPath()
        openPdf.close()
        moveEmailArchive(pdf_file, archivedPdfPath)

    print(extraction_info)

schedule.every(2).seconds.do(main)

while __name__ == '__main__':
    schedule.run_pending()
    time.sleep(1)
