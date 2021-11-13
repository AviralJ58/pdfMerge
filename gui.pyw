import os, PyPDF2, win32com.client
import tkinter
from tkinter import *
from tkinter import filedialog
import tkinter.messagebox

def display_files(location):
    global all_pdfs
    global converted
    global powerpoint
    global word
    global rows
    rows=8
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
                l=Label(root,text=f'{count}. {temp}')
                l.grid(sticky='w',row=rows,column=0)
                rows+=1
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
                l=Label(root,text=f'{count}. {temp}')
                l.grid(sticky='w',row=rows,column=0)
                rows+=1
                count+=1

            else:
                all_pdfs.append(filename)
                l=Label(root,text=f'{count}. {filename}')
                l.grid(sticky='w',row=rows,column=0)
                rows+=1
                count+=1

def select_files(flag):
    global all_pdfs
    global to_merge
    global rows
    global e
    global msg

    if (flag=='Y'):
        to_merge = all_pdfs

    elif (flag=='N'):
        to_merge = []
        global k
        req=list(map(int,k.get().split(',')))
        for i in req:
            to_merge.append(all_pdfs[i-1])

    rows+=1
    msg=Label(root,text="Merging PDFs. Please wait!",bg="yellow",width=80).grid(row=rows)
    merge_pdfs(e.get())    

def merge_pdfs(output):
    global converted
    global msg

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

    msg=Label(root,text="PDFs merged!",bg="green",width=80).grid(row=rows)

def compress_pdf(inp):
    global msg

    # tkinter.messagebox(root,text="Compressing PDF. Please wait!",bg="yellow",width=80).grid(row=rows)
    try:
        from pylovepdf.ilovepdf import ILovePdf
        public_key='project_public_48f3e3103bf52723e23da3527de74647_EwXuCf35023e65ee10f2c620e18f17f215125'
        ilovepdf=ILovePdf(public_key, verify_ssl=True)
        compressor=ilovepdf.new_task('compress')
        compressor.add_file(inp+'.pdf')
        compressor.set_output_folder(location)
        compressor.execute()
        compressor.download()
        compressor.delete_current_task()
        os.remove(e.get()+'.pdf')

        msg=Label(root,text="PDF compressed!",bg="green",width=100).grid(row=rows)

    except:
        msg=Label(root,text="Can't compress the file as the connection can't be established due to a network error.",bg="red",width=80).grid(row=rows)

root=Tk()
root.title("pdfMerge")

location = filedialog.askdirectory(initialdir="/",title="Select a Directory")

lab=Label(root,text="File Directory:  "+location,fg="black")
lab.grid(sticky='w',row=0,column=0)
display_files(location)

lab=Label(root,text='*-*-*-*-*-*-*-FILES IN ORDER*-*-*-*-*-*-*-*-*').grid(sticky='w',column=0,row=5)

r = StringVar()
Label(root,text="Want to merge all files?:").grid(sticky='w',column=0,row=1)
Radiobutton(root,text="Yes",variable=r,value="Y").grid(sticky='w',row=1,column=1)
Radiobutton(root,text="No",variable=r,value="N").grid(sticky='w',row=1,column=2)

Label(root,text="Output Filename:").grid(sticky='w',column=0,row=2)
e=Entry(root,text="filename",width=50,borderwidth=5)
e.grid(sticky='w',row=2,column=1)

Label(root,text="If all files are not to be merged, enter file no. in order separated by ',' :").grid(sticky='w',column=0,row=3)
k=Entry(root,text="fileno",width=50,borderwidth=5)
k.grid(sticky='w',row=3,column=1)

Button(text="Merge",bg="grey",fg="white",command=lambda:select_files(r.get())).grid(sticky='w',column=1,row=4)

Button(text="Compress",bg="grey",fg="white",command=lambda:compress_pdf(e.get())).grid(sticky='w',column=1,row=5)

root.mainloop() 