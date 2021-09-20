import os, PyPDF2, win32com.client

location=input("Folder path:\n")

os.chdir(location)

output=input("Output Filename:\n")

powerpoint = win32com.client.Dispatch("Powerpoint.Application")

count=1
all_pdfs = []
converted = []
print("The PDFs in the folder are: ")

for filename in os.listdir("."):

    if (filename.endswith('.pdf') or filename.endswith('.ppt') or filename.endswith('.pptx')):

        if (filename.endswith('.ppt') or filename.endswith('.pptx')):
            temp=filename.split('.')[0]
            temp=temp+'.pdf'
            inputfile=os.path.abspath(filename)
            outputfile=os.path.abspath(temp)
            ppt = powerpoint.Presentations.Open(inputfile)
            ppt.SaveAs(outputfile, 32)
            ppt.Close()
            
            all_pdfs.append(temp)
            converted.append(temp)
            print(f'{count}. {temp}')
            count+=1
        
        else:
            all_pdfs.append(filename)
            print(f'{count}. {filename}')
            count+=1


flag=input('Merge all files in displayed order? Y/N\n')

if (flag.upper()=='Y'):
    to_merge = all_pdfs

elif (flag.upper()=='N'):
    to_merge = []
    print("Enter the respective number of PDFs to be merged according to the required order, separated by comma:")
    req=list(map(int,input().split(',')))
    for i in req:
        to_merge.append(all_pdfs[i-1])

writer = PyPDF2.PdfFileWriter()
pdfOutput = open(output+".pdf","wb")

for filename in to_merge:
    req_file = open(filename,"rb")
    reader = PyPDF2.PdfFileReader(req_file)

    for pgNo in range(reader.numPages):
        current_page = reader.getPage(pgNo)
        writer.addPage(current_page)
        
    writer.write(pdfOutput)         
    req_file.close()
        
pdfOutput.close()

if len(converted)>0:
    powerpoint.Quit()
    for filename in converted:
        os.remove(filename)

wait=input('PDFs merged! Check the source folder for output file. Press ENTER to exit.')