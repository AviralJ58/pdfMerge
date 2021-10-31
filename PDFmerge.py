import os, PyPDF2, win32com.client

def display_files(location):
    global all_pdfs
    global converted
    global powerpoint
    global word
    os.chdir(location)
    all_pdfs = []
    converted = []
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    word = win32com.client.Dispatch("Word.Application")
    count=1

    for filename in os.listdir("."):

        if (filename.endswith('.pdf') or filename.endswith('.ppt') or filename.endswith('.pptx') or filename.endswith('.doc') or filename.endswith('.docx')):

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
            
            elif (filename.endswith('.doc') or filename.endswith('.docx')):
                temp=filename.split('.')[0]
                temp=temp+'.pdf'
                inputfile=os.path.abspath(filename)
                outputfile=os.path.abspath(temp)
                doc = word.Documents.Open(inputfile)
                doc.SaveAs(outputfile, 17)
                doc.Close()
                
                all_pdfs.append(temp)
                converted.append(temp)
                print(f'{count}. {temp}')
                count+=1

            else:
                all_pdfs.append(filename)
                print(f'{count}. {filename}')
                count+=1

def select_files(flag):
    global all_pdfs
    global to_merge

    if (flag.upper()=='Y'):
        to_merge = all_pdfs

    elif (flag.upper()=='N'):
        to_merge = []
        print("Enter the respective number of PDFs to be merged according to the required order, separated by comma:")
        req=list(map(int,input().split(',')))
        for i in req:
            to_merge.append(all_pdfs[i-1])

def merge_pdfs(output):
    global converted
    
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
        word.Quit()
        for filename in converted:
            os.remove(filename)

location=input("Folder path:\n")
print("The PDFs in the folder are: ")
display_files(location)

flag=input('Merge all files in displayed order? Y/N\n')
select_files(flag)

output=input("Output Filename:\n")
merge_pdfs(output)

wait=input('PDFs merged! Check the source folder for output file. Press ENTER to exit.')