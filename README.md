# PDFmerge

GUI based application to merge the **PDF/PPT/PPTX/DOC/DOCX** in the specified folder with given output filename.

If the folder has PPT/PPTX/DOC/DOCX then it will be automatically converted to PDF (temporary files) and can be merged. The converted PDFs will be deleted from the directory when the process is completed. After merging, the user will be given an option to compress the final merged PDF.

>⚠️ **Only supports Windows OS**

**Modules used**: PyPDF2, win32com, pylovepdf, tkinter

The user has **two options**:
 1) Merge all the files in the directory.
 2) Merge selective files by choosing from a list of files available in the directory in specified order.

The user can choose to compress the final PDF file. However, this requires an active internet connection. 