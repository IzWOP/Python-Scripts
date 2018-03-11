import PyPDF2
import re
import xlsxwriter

docsFile = open('image0001.pdf','rb')
pdfReader = PyPDF2.PdfFileReader(docsFile)
loanNumberlist = []
loan2Matchlist = []
poolNumlist = []
borrowerNamelist = []
wb = xlsxwriter.Workbook('docInfo.xlsx')
ws = wb.add_worksheet('sheet2')
row = 0
#setting column Headers
columnHeaders = ['Borrower Name', 'Loan Number', 'LD Loan Number', 'Pool #']
for col, colname in enumerate(columnHeaders, start=0):
    ws.write(row, col, colname)


class pdfExtract:
    def __init__(self, pg):
        self.pg = pg
    def extractShit(self):
        #grab page from for loop
        pageObj = pdfReader.getpage(self.pg) 
        #page text to variable
        pgData = pageObj.extractText()
        #find loan number and append to loanNumberlist
        loanNumber = re.split('\\bLoan #:\\b', pgData)[-1]
        loanNumberlist.append(loanNumber)
        #find data after / and append
        loan2Match = re.match(r"?:/\d{0,10}", pgData)[-1]
        loan2Matchlist.append(loan2Match)
        #grab pool number and append
        poolNumber = re.split('\\bPool #:\\b',pgData)[-1]
        poolNumlist.append(poolNumber)
        #find borrower name and append
        borrowerName =re.split('\\bBorrower #:\\b',pgData)[-1]
        borrowerNamelist.append(borrowerName)
#going through all pages
for page in range(0, 223):
    pdfExtract(page) 
#grabbing finished list to populate every row on first column
for row, rowvar in enumerate(borrowerNamelist, start=1):#write Borrower name
    col = 0
    ws.write(row, col, rowvar)
#adding info from second list to second column
for row, lnNM in enumerate(loanNumberlist, start=1):#write loan number 1
    col = 1
    ws.write(row, col, lnNM)
#adding info to third column
for row, lnNM2 in enumerate(loan2Matchlist, start=1):#write loan number 2
    col = 2
    ws.write(row, col, lnNM2)
#adding data to last column
for row, plNm in enumerate(poolNumlist, start=1):#write pool number
    col = 3
    ws.write(row, col, plNm)
#close workbook
wb.close()