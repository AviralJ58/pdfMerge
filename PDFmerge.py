import os, PyPDF2

location=input("Folder path:\n")

os.chdir(location)

output=input("Output filename:\n")

number=input('Enter number of PDFs to be merged. Enter "all" to merge all PDFs in the folder.\n')

if (number=='all'):
    to_merge = []
    for filename in os.listdir("."):
        if filename.endswith('.pdf'):
            to_merge.append(filename)

else:
    number=int(number)
    to_merge = []
    print("Enter the names of PDF to be merged (with extension):")
    for i in range (number):
        filename=input(f'{i+1}: ')
        if filename.endswith('.pdf'):
            to_merge.append(filename)

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