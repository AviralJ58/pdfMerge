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
    scrollbar = Scrollbar(frm, orient="vertical")
    scrollbar.pack(side=RIGHT, fill=Y)
    listNodes = Listbox(frm, width=60, yscrollcommand=scrollbar.set,background='#282C34',fg='white',selectbackground='#282C34',selectforeground='white')
    scrollbar.config(command=listNodes.yview)
    listNodes.pack(expand=True, fill=Y)

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
                s=str(count)+'.'+(temp)
                listNodes.insert(END, s)
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
                s=str(count)+'.'+(temp)
                listNodes.insert(END, s)
                count+=1

            else:
                all_pdfs.append(filename)
                s=str(count)+'.'+(filename)
                listNodes.insert(END, s)
                count+=1
        rows+=1

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
            to_merge.append(all_pdfs[i - 1])
    rows+=1

    root.geometry("600x582")
    msg=Label(root,text="Merging PDFs",bg="yellow",width=100).grid(row=rows,columnspan=3)
    tkinter.messagebox.showinfo("Processing","Please wait for the process to complete!")
    merge_pdfs(e.get())

def merge_pdfs(output):
    global converted
    global msg

    try:
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

        msg=Label(root,text="PDFs merged!",bg="green",width=100).grid(row=rows,columnspan=3)
        tkinter.messagebox.showinfo("Files Merged","Check the source folder for merged PDF!")

    except:
        msg=Label(root,text="PDFs merged!",bg="green",width=100).grid(row=rows,columnspan=3)
        tkinter.messagebox.showinfo("Files Merged","A few files had problems, but they are merged. Check the source folder for merged PDF!")

    finally:
        button4.config(state=NORMAL)
        if len(converted)>0:
            powerpoint.Quit()
            word.Quit()
            for filename in converted:
                os.remove(filename)

def compressConf():
    msg=Label(root,text="Compressing PDFs. Please wait!",bg="yellow",width=100).grid(row=rows,columnspan=3)
    res=tkinter.messagebox.askyesno("Confirmation","Compression may take some time. Do you wish to continue?")
    if res:
        compress_pdf(e.get())
    else:
        tkinter.messagebox.showerror("Cancelled","Check the source folder for merged PDF!")
        msg=Label(root,text="Compression aborted by user.",bg="red",width=100).grid(row=rows,columnspan=3)


def compress_pdf(inp):
    global msg,filename

    from pylovepdf.ilovepdf import ILovePdf
    try:
        public_key='project_public_48f3e3103bf52723e23da3527de74647_EwXuCf35023e65ee10f2c620e18f17f215125'
        ilovepdf=ILovePdf(public_key, verify_ssl=True)
        compressor=ilovepdf.new_task('compress')
        compressor.add_file(inp+'.pdf')
        compressor.set_output_folder(filename)
        compressor.execute()
        compressor.download()
        compressor.delete_current_task()
        os.remove(e.get()+'.pdf')

        msg=Label(root,text="PDF compressed!",bg="green",width=100).grid(row=rows,columnspan=3)
        tkinter.messagebox.showinfo("Success","Check the source folder for compressed PDF!")

    except:
        msg=Label(root,text="Can't compress the file as the connection can't be established.",bg="red",width=100).grid(row=rows,columnspan=3)
        tkinter.messagebox.showerror("Error","Can't compress. Check the source folder for merged PDF!")

def browse_button():
    global folder_path,filename,button2,button3
    filename = filedialog.askdirectory()
    folder_path.set(filename)
    lab=Label(root,bg='#282C34',fg='white',text='FILES IN ORDER').grid(column=0,row=6,columnspan=3,sticky=W)
    display_files(filename)
    button2.config(state=DISABLED)
    button3.config(state=NORMAL)
    root.geometry("600x562")

root=Tk()
root.title("pdfMerge")
root.configure(bg='#282C34')
root.geometry('600x400')
photo = PhotoImage(file = './pdfmerge-icon.png')
root.iconbitmap(r'pdfmerge-icon.ico')
root.iconphoto(True, photo)
frm = Frame(root)
frm.grid(row=7, column=0, sticky=N+S)
root.rowconfigure(1, weight=1)
root.columnconfigure(1, weight=1)
root.resizable(False,False)
#add image to the top
photo = PhotoImage(file = './pdfmerge-icon.png')
label = Label(root, image=photo,bg='#282C34')
label.grid(row=0, column=0, sticky=W+E+N+S,columnspan=3)

lab=Label(root,bg='#282C34',fg='white',text="File Directory:  ")
lab.grid(row=1,column=0)
folder_path=StringVar()
lbl1 = Label(master=root,bg='#282C34',fg='white',textvariable=folder_path)
lbl1.grid(row=1, column=1)
button2 = Button(text="Browse", command=browse_button,fg='white',bg='grey')
button2.grid(row=1, column=2, padx=10, sticky=E+W+N)

r = StringVar()
Label(root,bg='#282C34',fg='white',text="Want to merge all files?").grid(column=0,row=2,sticky=E+W+N)
Radiobutton(root,text="Yes",variable=r,value="Y",bg='#282C34',fg='green').grid(row=2,column=1, sticky=E+W+N)
Radiobutton(root,text="No",variable=r,value="N",bg='#282C34',fg='red').grid(row=2,column=2, sticky=E+W+N)

Label(root,bg='#282C34',fg='white',text="Output Filename:").grid(column=0,row=3)
e=Entry(root,text="filename",width=40,borderwidth=5)
e.grid(row=3,column=1,columnspan=3, sticky=E+W+N, padx=10)

Label(root,bg='#282C34',fg='white',text="If all files are not to be merged, enter file no. in order separated by ',' :").grid(column=0,row=4, sticky=E+W+N, pady=10)
k=Entry(root,text="fileno",width=40,borderwidth=5,state=DISABLED)
k.grid(row=4,column=1, pady=10,columnspan=3,sticky=E+W+N, padx=10)

button3=Button(text="Merge",bg="grey",fg="white",command=lambda:select_files(r.get()),state='disabled')
button3.grid(column=1,row=5, sticky=E+W+N, padx=10, pady=10)
button4=Button(text="Compress",bg="grey",fg="white",command=lambda:compressConf(),state='disabled')
button4.grid(column=2,row=5, sticky=E+W+N, padx=10, pady=10, ipadx=13, ipady=1)

root.mainloop()
