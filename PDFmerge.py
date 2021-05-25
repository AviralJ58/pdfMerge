import os, PyPDF2

location=input("Folder path:\n")

os.chdir(location)

output=input("Output filename:\n")

number=input('Enter number of PDFs to be merged. Enter "all" to merge all PDFs in the folder.\n')

all_pdfs = []
for filename in os.listdir("."):
    if filename.endswith('.pdf'):
        all_pdfs.append(filename)

if (number=='all'):
    to_merge = all_pdfs

else:
    number=int(number)
    print("PDFs in folder:")
    for i in range(len(all_pdfs)):
        print(f'{i+1}. {all_pdfs[i]}')
    to_merge = []
    print("Enter the respective number of pdfs to be merged separated by comma:")
    req=list(map(int,input().split(',')))
    for i in req:
        to_merge.append(all_pdfs[i-1])

writer = PyPDF2.PdfFileWriter()

for filename in to_merge:
    req_file = open(filename,"rb")
    reader = PyPDF2.PdfFileReader(req_file)
    for pgNo in range(reader.numPages):
        current_page = reader.getPage(pgNo)
        writer.addPage(current_page)
        
pdfOutput = open(output+".pdf","wb")
writer.write(pdfOutput)
pdfOutput.close()
wait=input('PDFs merged! Check the source folder for output file. Press ENTER to exit.')