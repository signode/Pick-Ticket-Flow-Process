'''
IMPROVMENTS:

1 - MAKE COUNTER() FOR 'PAGEITER' TO CAPTURE DUPLICATES
    PROBLEM SOLVED - ANY DUPLICATES WILL BE APPENDED ONE AFTER ANOTHER TOGETHER

PROBLEMS:

1 - IF PAGES WITH SAME ORDERNO ARE SCATTERED WITHIN THE SX GENERATED PDF, THEN PROGRAM APPENDS 'FIRST FOUND + LAST FOUND' PAGE ONLY (MULTIPAGE PAGE ORDERNO'S MUST BE CONSISTANT)
2 - IF PAGES WITH SAME ORDERNO ARE FOUND ACROSS SX GENERATED PDF'S, A CMD MSSG IS GENRATED FOR SUCH DUPLICATES, BUT ONLY THE LATEST PDF WILL BE EXTRACTED AND/OR MERGERED  

TO START THE PROGRAM : UNCOMMENT THE __name__ == __main__ condition

'''


''' Imports '''
try:
    import re2 as re
except ImportError:
    import re
from PyPDF2 import PdfFileWriter, PdfFileMerger, PdfFileReader
from datetime import date
import glob 
import os
from prettytable import PrettyTable


def main():

    def listAllPdfFiles(path):
        return glob.glob(os.path.join(path,'*.pdf'))

    today = date.today()
    extraction_info = PrettyTable()
    extraction_info.field_names = ["pdf file name", "num of pdfs extracted", "multiple page pickTicket"]

    
    # ogpath = 'C:\\Users\\0235124\\OneDrive - University of Waterloo\\Desktop\\signodeProjects\\pdfParse\\pyPDF2\\OneDrive\\'
    path_for_SX_attachments = 'OneDrive'
    ogpath = os.getcwd() + "\\" + path_for_SX_attachments + '\\'
    path = ogpath + str(today)
    emailArchivePath = ogpath + 'email_archive'
    pdf_files = listAllPdfFiles(ogpath)

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
                print(f"Can't change the Current Working Directory (Check if the dir was created) to {path}")        
                print("Current Working Directory " , os.getcwd())

    def returnToOgPath() : os.chdir(ogpath)

    def prettyInfo(pdf, pages, multiPage, temp):
        temp = temp.copy()
        temp = list(temp)
        temp_str = ''
        for _ in temp:
            temp_str += ' ' + _
        if multiPage:
            out = 'Found' + temp_str
        else:
            out = ''
        extraction_info.add_row([os.path.basename(pdf), pages, out])
    
    def pdfReader(file) : return PdfFileReader(file)

    def save_Rename_Pdf(pages, file, multiPagePickTickets, duplicatePdfs):   
        uniquePickTickets = set()
        setToPrint = set() # console output (multipage order no for every SX pdf)
        multiPageValid = False
        # pageIter = 1 # used to locate multipage pickTickets
        # duplicatePdfs = [] # Stores Duplicates 
        for page in range(pages):
            pageInfo = inputPdf.getPage(page)
            rawData = pageInfo.extractText() # extract initial data
            rawData = re.sub(r'\r\n', ' ', rawData) # substitute \n with ' '
            rawData = pickTicketRawInfo(rawData)
            pickTicketNo = orderNo(rawData[0].group(0)).group(0)
            orders.append(pickTicketNo)

            output = PdfFileWriter()
            output.addPage(pageInfo)
            
            if pickTicketNo in uniquePickTickets:
                # print(uniquePickTickets)
                setToPrint.add(pickTicketNo)
                pageIter += 1
                # print(pickTicketNo, pageIter)
                if pageIter > 1:
                    multiPageValid = True
                multiPagePickTickets.add(pickTicketNo)
                with open(f"{pickTicketNo} ({pageIter}).pdf", 'wb') as outputStream:
                    output.write(outputStream)
            else:
                pageIter = 1
                if pickTicketNo in multiPagePickTickets:
                    duplicatePdfs.append(pickTicketNo)
                # else:
                #     pageIter = 1
                with open(f"{pickTicketNo}.pdf", "wb") as outputStream: # create a orderNo list to extract each # (orderNo(temp[i]).group(0))
                    output.write(outputStream)
                    uniquePickTickets.add(pickTicketNo)


        prettyInfo(file, pages, multiPageValid, setToPrint)


    def moveEmailArchive(src, dst):
        # Path(src).rename(dst)
        os.replace(src, dst) 

    def rename(oldName, newName):
        os.rename(oldName, newName)

    def mergePickTickets(targets):
        for ticket in targets:
            old_name = path + '\\' + ticket + '.pdf'
            new_name = path + '\\' + ticket + ' (1).pdf'
            if os.path.exists(old_name) and not os.path.exists(new_name):
                rename(old_name, new_name)

            pdf_files = listAllPdfFiles(path)

            merge_these = []

            for file in pdf_files:
                if ticket in file:
                    merge_these.append(file)

            merger = PdfFileMerger()

            for pdf in merge_these:
                merger.append(pdf)

            os.chdir(path)
            merger.write(f"{ticket}.pdf")
            os.chdir(ogpath)
            merger.close()

            for file in pdf_files:
                if ticket in file:
                    if os.path.exists(file):
                        os.remove(file)         

    ''' Main Loop '''  
    def mainProgram():
        if not (pathCheck and emailPathCheck):
            createDir(path)
            createDir(emailArchivePath)

        enterDir(path)
        save_Rename_Pdf(noPages, pdf_file, toBeMerged, duplicates)
        returnToOgPath()
        openPdf.close()
        moveEmailArchive(pdf_file, archivedPdfPath)        

    toBeMerged = set()
    duplicates = []
    for pdf_file in pdf_files:
        openPdf = open(pdf_file, "rb")
        inputPdf = pdfReader(openPdf) 
        noPages = inputPdf.numPages

        orders = []

        archivedPdfPath = emailArchivePath + '\\' + os.path.basename(pdf_file) 

        pathCheck = os.path.exists(path)
        emailPathCheck = os.path.exists(emailArchivePath)

        mainProgram()
    
    # Merge MultiPage Pick Tickets
    toBeMerged = list(toBeMerged)
    mergePickTickets(toBeMerged)

    # Console Output
    if toBeMerged:
        print('Pick Tickets having more than one page:', toBeMerged)
    if duplicates:
        print('Pick Tickets Duplicated', str(list(set(duplicates))))
    print(extraction_info)

if __name__ == '__main__':
    main()