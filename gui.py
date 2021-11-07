import os, PyPDF2, win32com.client
import tkinter
from tkinter import *
from tkinter import filedialog
import tkinter.messagebox

rows=8

def display_files(location):
    global all_pdfs
    global converted
    global powerpoint
    global word
    global rows
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
                #print(f'{count}. {temp}')
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
                #print(f'{count}. {temp}')
                l=Label(root,text=f'{count}. {temp}')
                l.grid(sticky='w',row=rows,column=0)
                rows+=1
                count+=1

            else:
                all_pdfs.append(filename)
                #print(f'{count}. {filename}')
                l=Label(root,text=f'{count}. {filename}')
                l.grid(sticky='w',row=rows,column=0)
                rows+=1
                count+=1


def select_files(flag):
    global all_pdfs
    global to_merge
    global rows
    global e

    if (flag=='Y'):
        to_merge = all_pdfs

    elif (flag=='N'):
        to_merge = []
        global k
        req=list(map(int,k.get().split(',')))
        for i in req:
            to_merge.append(all_pdfs[i-1])

    Label(root,text=to_merge,bg="green").grid(row=rows)
    rows+=1
    merge_pdfs(e.get())
    

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

    # if len(converted)>0:
    #     powerpoint.Quit()
    #     word.Quit()
    #     for filename in converted:
    #         os.remove(filename)

def compress_pdf():
    from pylovepdf.ilovepdf import ILovePdf
    public_key='project_public_48f3e3103bf52723e23da3527de74647_EwXuCf35023e65ee10f2c620e18f17f215125'
    ilovepdf=ILovePdf(public_key, verify_ssl=True)
    compressor=ilovepdf.new_task('compress')
    compressor.add_file(inp+'.pdf')
    compressor.set_output_folder(location)
    compressor.execute()
    compressor.download()
    compressor.delete_current_task()

def showData():
    global rows
    Label(root,text=to_merge).grid(row=rows)
    rows+=1
    


root=Tk()
root.title("pdfMerge")

location = filedialog.askdirectory(initialdir="/",title="Select a Directory")

tkinter.messagebox.showinfo("ALERT","Please minimise the main window and reopen it! sorry for the technical issue!")

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
#Button(text="Merge",bg="grey",fg="white",command=lambda:showData()).grid(sticky='w',column=1,row=4)



root.mainloop() 



# if (input("PDFs merged! Do you want to compress the pdf (requires internet)? Y/N\n").upper()=='Y'):
#     try:
#         compress_pdf(output)
#         os.remove(output+'.pdf')
#         wait=input('PDF compressed! Check the source folder for output file. Press ENTER to exit.')
#     except:
#         wait=input("Can't compress the file as the connection can't be established due to a network error. Check the source folder for output file. Press ENTER to exit.")

# else:
#     wait=input('Process complete! Check the source folder for output file. Press ENTER to exit.')
