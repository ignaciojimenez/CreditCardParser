import re
import time
import sys
import os
import zipfile
import getopt
import shutil
import io
# Installed modules
import xlrd
from PyPDF2 import PdfFileReader


def checksum(string):
    '''
    Check the Luhn algorithm on the given string
    Returns True/False
    '''
    odd_sum = sum(list(map(int, string))[-1::-2])
    even_sum = sum([sum(divmod(2 * d, 10)) for d in list(map(int, string))[-2::-2]])
    return ((odd_sum + even_sum) % 10 == 0)


def searchInLine(textLine, regex_list):
    '''
    Looks for credit cards within a given plain text line
    Prints the credit card match if it occurs
    Retuns 1 if a CC is found 0 if not
    '''
    # TODO Search for multiple matches within single line
    for regEx in regex_list:
        m = re.search(r"%s" % regEx[1].rstrip(), textLine)
        if m:
            if checksum(m.group(0)):
                return regEx[0]
            else:
                return None


def textFSearch(cc_path, regex_list):
    '''
    Looks for credit cards within a given plain text file
    '''
    foundtxt = ""
    txt_cc_list = dict([(row[0], 0) for row in regex_list])

    with open(cc_path, 'r', encoding="latin-1") as cc_file_data:
        for cc_file_line in cc_file_data:
            foundtxt = searchInLine(cc_file_line, regex_list)
            if foundtxt:
                txt_cc_list[foundtxt] += 1
    return list(txt_cc_list.values())


def pdfFSearch(cc_path, regex_list):
    '''
    Looks for credit cards within a given Excel File
    Retuns int total of credit cards found
    '''
    pdfPageCount = 0  # Page iterator counter
    text = ""  # Contains all the extracted text
    pdf_cc_list = dict([(row[0], 0) for row in regex_list])
    pdfReader = PdfFileReader(open(cc_path, 'rb'))
    # While loop will read each page
    while pdfPageCount < pdfReader.numPages:
        pageObj = pdfReader.getPage(pdfPageCount)
        text += pageObj.extractText()
        buf = io.StringIO(pageObj.extractText())
        for line in buf:
            foundpdf = searchInLine(line, regex_list)
            if foundpdf:
                pdf_cc_list[foundpdf] += 1
        pdfPageCount += 1
    # If everything included in the PDF is scanned (PyPDF cannot extract text from images).
    if text == "":
        pdf_cc_list = dict.fromkeys(pdf_cc_list, None)
    return list(pdf_cc_list.values())


def excelFSearch(cc_path, regex_list):
    '''
    Looks for credit cards within a given Excel File
    Retuns int total of credit cards found
    '''
    foundxls = 0
    xls_cc_list = dict([(row[0], 0) for row in regex_list])
    with xlrd.open_workbook(cc_path) as wb:
        for sheet in wb.sheets():
            for row in range(sheet.nrows):
                foundxls = searchInLine(','.join(sheet.row_values(row)), regex_list)
                if foundxls:
                    xls_cc_list[foundxls] += 1
    return list(xls_cc_list.values())


def zipFSearch(cc_path, regex_list):
    '''
    Looks for credit cards within a given zip File
    Retuns int total of credit cards found
    '''
    # TODO Print relative paths not absolute
    with zipfile.ZipFile(cc_path, 'r') as zfile:
        # print("-- Opening ZIP FILE --> " + cc_path)
        for finfo in zfile.infolist():
            list_fichero_zip = searchInFile(zfile.extract(finfo), regex_list)
            if not list_fichero_zip:
                pass
            else:
                print(cc_path + "\\" + finfo.filename + "," + str(list_fichero_zip).strip('[]'))

            if "/" in finfo.filename:
                shutil.rmtree(finfo.filename.split("/")[0])
            else:
                os.remove(finfo.filename)
    # print("-- End of ZIP FILE --> " + cc_path)
    return None


def searchInFile(cc_path, regex_list):
    '''
    Looks for credit cards within a given file
    Retuns int total of credit cards found, -1 if file is not supported
    '''
    if any(cc_path[-3:].lower() in s for s in unsupported_files):
        return None
    elif any(cc_path[-3:].lower() in s for s in ['xls', 'xlsx']):
        return excelFSearch(cc_path, regex_list)
    elif any(cc_path[-3:].lower() in s for s in ['pdf']):
        return pdfFSearch(cc_path, regex_list)
    elif zipfile.is_zipfile(cc_path):
        return zipFSearch(cc_path, regex_list)
    else:
        # TODO Auto-find out if file can be read as plain text
        return textFSearch(cc_path, regex_list)


def searchInDir(cc_path, regex_list):
    '''
    Searchs for credit cards in files contained within a given path
    Returns tuple containing Files read, Files analyzed, Files including credit cards, Total credit cards found
    '''
    # Loop over the given path
    for root, dirs, files in os.walk(cc_path):
        for filename in files:
            list_fichero = searchInFile(os.path.join(root, filename), regex_list)
            print(os.path.join(root, filename) + "," + str(list_fichero).strip('[]'))


if __name__ == '__main__':
    # reading input options
    # reading input options
    inputfile = ''
    inputdir = ''

    try:
        opts, args = getopt.getopt(sys.argv[1:], "hi:d:")
    except getopt.GetoptError:
        # TODO Read current dir files as the default option
        print('Syntax: CreditCardSearch.py -i <inputfile> | -d <inputdirectory>')
        sys.exit(2)
    if not opts:
        print('Syntax: CreditCardSearch.py -i <inputfile> | -d <inputdirectory>')
        sys.exit(2)

    for opt, arg in opts:
        if opt == '-h':
            print('Usage CreditCardSearch.py -i <inputfile> -d <inputdirectory> | -o <outputfile>')
            sys.exit()
        elif opt in ("-i"):
            inputfile = arg
        elif opt in ("-d"):
            inputdir = arg

    # File extension support variables
    # TODO PDF, docx, pptx, etc.
    tested_files = ['txt', 'csv', 'xls', 'xlsx', 'rtf', 'xml', 'html', 'json', 'zip', 'pdf']
    unsupported_files = ['doc', 'docx', 'pptx', 'jpg', 'gif',
                         'png', 'mp3', 'mp4', 'wav', 'aiff', 'mkv', 'avi', 'exe', 'dll']

    # Print memo when the script starts
    print("[" + time.ctime() + "] CreditCardSearch started")

    # Read CSV regex file to be tested
    regex_path = 'regexcard.csv'
    regex_list = []
    with open(regex_path, 'r') as regex_file:
        for line in regex_file:
            line_list = line.rstrip().split(',')
            regex_list.append(line_list)

    print("file," + ', '.join([row[0] for row in regex_list]))
    if inputdir:
        searchInDir(inputdir, regex_list)
    elif inputfile:
        print(os.path.join(inputfile) + "," + str(searchInFile(inputfile, regex_list)).strip('[]'))
