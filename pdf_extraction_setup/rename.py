import os
import glob
import re

ogpath = 'C:\\Users\\0235124\\Desktop\\signodeProjects\\pdfParse\\pyPDF2'
pdf_files = glob.glob(os.path.join(ogpath,'*.pdf'))

temp = []
for pdf in pdf_files:
    # temp.append(os.path.basename(pdf))
    pattern = re.compile(r'([(m | 2)](.*)pdf$)')
    patternMatches = pattern.finditer(os.path.basename(pdf))
    pickTicketData = [match for match in patternMatches]
    # temp.append(pickTicketData[0].group(0))
    temp_path = ogpath + '//' + pickTicketData[0].group(0)
    os.rename(pdf, temp_path)

# print(temp)