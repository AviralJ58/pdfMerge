import os, PyPDF2

location=input("Folder path:\n")

os.chdir(location)

output=input("Output Filename:\n")

count=1
all_pdfs = []
print("The PDFs in the folder are: ")
for filename in os.listdir("."):
    if filename.endswith('.pdf'):
        all_pdfs.append(filename)
        print(f'{count}. {filename}')
        count+=1

flag=input('Merge all PDFs in displayed order? Y/N\n')

if (flag.upper()=='Y'):
    to_merge = all_pdfs

elif (flag.upper()=='N'):
    to_merge = []
    print("Enter the respective number of PDFs to be merged according to the required order, separated by comma:")
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