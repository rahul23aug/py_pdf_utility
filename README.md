# py_pdf_utility
This is a simple utility designed using pyqt5 it can merge, convert and split PDF files

Features:
- Automatically crawl folders for relevant files
- Automatically convert to pdf while merging
- convert only mode
- split pdf files

UI explained
1- Folder / File path. copy/paste the path, if it is a folder it will automatically crawl folders

2- Select individual files

3- Change where the processed files should be saved default is desktop (ensure other locations allow access)

4- Load folder to select relevant files within a folder

5- Merge pdf files, when selected it will convert image files to pdf and then merge

6- convert only mode

Build

To create an executable for this file using pyinstaller 
execute:

pyinstaller --windowed --onefile <path-to-PDFU.pyw>

