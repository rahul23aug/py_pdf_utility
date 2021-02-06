#---------------------------------------------------------------------
'''
PDF Utilities
Author:- Rahul V 
github:- rahul23aug
date: 2-6-2020

A simple program to merge/split and convert to pdfs 
'''
#---------------------------------------------------------------------
import os
try:
    from PyQt5 import QtCore, QtGui, QtWidgets
except:
    os.system('pip install pyqt5-tools')
    from PyQt5 import QtCore, QtGui, QtWidgets

try:
    from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
except:
    os.system('pip install PyPDF2')
    from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger

try:
    from PIL import Image
except:
    os.system('pip install Pillow')
    from PIL import Image
try:
    import img2pdf
except:
    os.system('pip install img2pdf')
    import img2pdf
try:
    import shutil
except:
    os.system('pip install shutil')
    import shutil
#--------------------------------------Global variables---------------------------------------------------------

file_names=[]
output_path=os.path.expanduser('~/Desktop').replace('\\','/') + '/PDFU/'

#output_path=os.path.abspath('./').replace('\\','/') + '/PDFU/'
print(output_path)
if not os.path.exists(output_path):
    os.mkdir(output_path)
if os.path.exists(output_path + "temp/"):
    shutil.rmtree(output_path + "temp/")
#Clean setup
if not os.path.exists(output_path + "temp/"):
    os.mkdir(output_path + "temp/")
progress_precentage = 0
convert_only = False
merge = False
default_message = ("DEBUG INFORMATION WINDOW\n" 
"PDF Utilities 1.0\n" 
"- Auto Convert to PDF before splitting\n" 
"- Can be used to convert only (select convert to PDF only).\n" 
"- To remove selected files from the list, simply double click it\n" 
"- If a follder cannot be accessed, run-as administrator\n" 
"- The list below gives the files that will be converted\n" 
"- Default Output Folder is Current Working Directory >> PDFU\n" 
"- Note: If needed set output folder during startup as it is reset during every rerun\n")
#---------------------------------------------------------------------------------------------------------------

#---------------------------------------------------Main Window-------------------------------------------------
class Ui_PDFU(object):
    def setupUi(self, PDFU):
        PDFU.setObjectName("PDFU")
        PDFU.resize(550, 471)
        PDFU.setMinimumSize(QtCore.QSize(550, 471))
        PDFU.setMaximumSize(QtCore.QSize(550, 471))
        self.centralwidget = QtWidgets.QWidget(PDFU)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.FileFolderPath = QtWidgets.QLineEdit(self.centralwidget)
        self.FileFolderPath.setObjectName("FileFolderPath")
        self.horizontalLayout.addWidget(self.FileFolderPath)
        self.FLoadButton = QtWidgets.QPushButton(self.centralwidget)
        self.FLoadButton.setObjectName("FLoadButton")
        self.horizontalLayout.addWidget(self.FLoadButton)
        self.FolderLoadButton = QtWidgets.QPushButton(self.centralwidget)
        self.FolderLoadButton.setObjectName("FolderLoadButton")
        self.horizontalLayout.addWidget(self.FolderLoadButton)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.verticalLayout.addWidget(self.line)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.OPFolder = QtWidgets.QLabel(self.centralwidget)
        self.OPFolder.setObjectName("OPFolder")
        self.horizontalLayout_2.addWidget(self.OPFolder)
        self.OpSetButton = QtWidgets.QPushButton(self.centralwidget)
        self.OpSetButton.setMaximumSize(QtCore.QSize(101, 31))
        self.OpSetButton.setObjectName("OpSetButton")
        self.horizontalLayout_2.addWidget(self.OpSetButton)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.verticalLayout.addWidget(self.plainTextEdit)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.Merge = QtWidgets.QCheckBox(self.centralwidget)
        self.Merge.setObjectName("Merge")
        self.horizontalLayout_3.addWidget(self.Merge)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setMaximumSize(QtCore.QSize(47, 13))
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_3.addWidget(self.label_3)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setMaximumSize(QtCore.QSize(241, 21))
        self.progressBar.setProperty("value", progress_precentage)
        self.progressBar.setObjectName("progressBar")
        self.horizontalLayout_3.addWidget(self.progressBar)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.listWidget = QtWidgets.QListWidget(self.centralwidget)
        self.listWidget.setObjectName("listWidget")
        self.verticalLayout.addWidget(self.listWidget)
        self.ConvertOnlyChkbx = QtWidgets.QCheckBox(self.centralwidget)
        self.ConvertOnlyChkbx.setObjectName("ConvertOnlyChkbx")
        self.verticalLayout.addWidget(self.ConvertOnlyChkbx)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.ClearAllSelButton = QtWidgets.QPushButton(self.centralwidget)
        self.ClearAllSelButton.setMaximumSize(QtCore.QSize(242, 23))
        self.ClearAllSelButton.setObjectName("ClearAllSelButton")
        self.horizontalLayout_4.addWidget(self.ClearAllSelButton)
        self.ProcessButton = QtWidgets.QPushButton(self.centralwidget)
        self.ProcessButton.setMaximumSize(QtCore.QSize(241, 23))
        self.ProcessButton.setObjectName("ProcessButton")
        self.horizontalLayout_4.addWidget(self.ProcessButton)
        self.verticalLayout.addLayout(self.horizontalLayout_4)
        self.verticalLayout_2.addLayout(self.verticalLayout)
        PDFU.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(PDFU)
        self.statusbar.setObjectName("statusbar")
        PDFU.setStatusBar(self.statusbar)
        self.menubar = QtWidgets.QMenuBar(PDFU)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 550, 21))
        self.menubar.setObjectName("menubar")
        PDFU.setMenuBar(self.menubar)
        self.retranslateUi(PDFU)
        QtCore.QMetaObject.connectSlotsByName(PDFU)
        self.ConvertOnlyChkbx.stateChanged.connect(self.convert_only)
        self.Merge.stateChanged.connect(self.merge_only)
        self.ClearAllSelButton.clicked.connect(self.state_reset)
        self.FLoadButton.clicked.connect(self.file_open_clicked)
        self.FolderLoadButton.clicked.connect(self.folder_open_clicked)
        self.FileFolderPath.textChanged.connect(self.checkValidPath)
        self.listWidget.itemDoubleClicked.connect(self.remove_list_item)
        self.listWidget.itemEntered.connect(self.validate_file_extensions)
        self.OpSetButton.clicked.connect(self.set_outfile_location)
        self.ProcessButton.clicked.connect(self.process_files)

    def retranslateUi(self, PDFU):
        _translate = QtCore.QCoreApplication.translate
        PDFU.setWindowTitle(_translate("PDFU", "PDF Utilities by Rahul V"))
        self.label.setText(_translate("PDFU", "Folder or File Path: "))
        self.FLoadButton.setText(_translate("PDFU", "Select File"))
        self.FolderLoadButton.setText(_translate("PDFU", "Load Folder"))
        self.OPFolder.setText(_translate("PDFU", "Current Output Folder: "+ output_path))
        self.OpSetButton.setText(_translate("PDFU", "Change"))
        self.Merge.setText(_translate("PDFU", "Merge"))
        self.plainTextEdit.setPlainText(_translate("PDFU", 
"PDF Utilities 1.0\n"
"- Auto Convert to PDF before splitting\n"
"- Can be used to convert only (select convert to PDF only).\n"
"- To remove selected files from the list, simply double click it\n"
"- If a follder cannot be accessed, run-as administrator\n"
"- The list below gives the files that will be converted\n"
"- Default Output Folder is Current Working Directory >> PDFU\n"
"- Note: If needed set output folder during startup as it is reset during every rerun\n"))
        self.label_3.setText(_translate("PDFU", "Progress: "))
        self.ConvertOnlyChkbx.setText(_translate("PDFU", "Convert to PDF only"))
        self.ClearAllSelButton.setText(_translate("PDFU", "Clear All Fields"))
        self.ProcessButton.setText(_translate("PDFU", "Convert and Split Files"))


    def set_outfile_location(self):
        global output_path
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        output_path = str(QtWidgets.QFileDialog.getExistingDirectory(None, "Select Folder")) + '/PDFU/'
        self.OPFolder.setText("Current Output Folder: "+ output_path)
        if not os.path.exists(output_path):
            os.mkdir(output_path)


    def state_reset(self):
        #Call this only after all tasks are done
        self.plainTextEdit.setPlainText("Tip: if you want to convert a file and not split it, check the box Convert Only\n")
        self.progressBar.setProperty("value", progress_precentage)
        global file_names
        file_names = []
        global merge
        merge = False
        global convert_only
        convert_only = False
        self.ConvertOnlyChkbx.setChecked(False)
        self.Merge.setChecked(False)
        self.listWidget.clear()
        global default_message
        self.update_banner(default_message)


    def checkValidPath(self): #check if the lineEdit field path is valid and processes accordingly
        global file_names
        if os.path.isdir(r'{}'.format(self.FileFolderPath.text())):
            self.crawl_folders_recursively(self.FileFolderPath.text())
            self.validate_file_extensions()
            self.add_to_list_widget()

        if os.path.isfile(self.FileFolderPath.text()):
            file_names.append(self.FileFolderPath.text())
            self.validate_file_extensions()
            self.add_to_list_widget()
        print(file_names)
        self.validate_file_extensions()

    def convert_only(self):
        global convert_only
        global merge
        global default_message
        if self.ConvertOnlyChkbx.isChecked():
            self.update_banner("[*]Convert Only Mode....\n\nFiltering non-image Files")
            self.ProcessButton.setText( "Convert Files")
            convert_only = True
            merge = False
            self.Merge.setChecked(False)
        else:
            convert_only = False
            self.update_banner("[*]Convert and Split Mode...\n\n"+default_message)
            self.ProcessButton.setText( "Convert and Split Files")

    def merge_only(self):
        global merge
        global convert_only
        global default_message
        if self.Merge.isChecked():
            merge = True
            convert_only = False
            self.ConvertOnlyChkbx.setChecked(False)
            self.update_banner("[*]Convert and Merge Mode....\n\nConverting non-pdf Files to PDF then Merging")
            self.ProcessButton.setText( "Convert and Merge Files")
            self.validate_file_extensions()
        else:
            merge = False
            self.update_banner("[*]Convert and Split Mode...\n\n"+default_message)
            self.ProcessButton.setText( "Process Files")

    def file_name_processor(self,file,opt='get_name'):
        if opt == 'get_ext':
            base=os.path.basename(file)
            return os.path.splitext(base)[-1]  #returns extension as .ext, using -1 ensures only the trailing .ext is kept
        else:
            base=os.path.basename(file)
            return os.path.splitext(base)[0]   #returns name of the file

    def crawl_folders_recursively(self,p):
        f =[]
        count = 0
        for root, dirs, files in os.walk(os.path.abspath(p)):
            global file_names
            for file in files:
                path_var=(os.path.join(root, file))
                f.append(path_var)
                f = list(set(f))
                #print(f)
        for i in f:
            if self.file_name_processor(i,'get_ext') in ['.jpg','.png','.jpeg','.pdf']:
                file_names.append(i)
            count+= 1
            self.update_banner("[i] Processing Files.........{} files processed!".format(count))
        print('\n[i] Folders crawled')

    def file_open_clicked(self):
        #open FileNamesDialog(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(None,"Select Files", "","All Files(*);;Image JPG (*.jpg);;Image JPEG (*.jpeg);;Image PNG (*.png);;PDF (*.pdf);; BMP (*.bmp);;TIFF (*.tiff) ", options=options)
        files = list(set(files))
        print(files)
        #iterate through filenames

        try:
            global file_names
            if files:
                for i in files :
                    file_names.append(i)
            s = set(file_names)
            file_names = list(s)
            # Avoid triggering SHGetFileInfo() timed out
            self.filter_shorcuts()
            self.validate_file_extensions()
            print(file_names)
            self.add_to_list_widget()
            for i in range(150):
                self.validate_file_extensions()
           #need to call validate twice for cleaner results
        except UnboundLocalError:
            pass


    def folder_open_clicked(self):
        self.update_banner("[i] Search Folder.....")	
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        folder = str(QtWidgets.QFileDialog.getExistingDirectory(None, "Select Directory")) #parent set to NONE
        try:
            if folder:
                print(folder)
                self.crawl_folders_recursively(folder)		
            self.filter_shorcuts()
            self.validate_file_extensions()
            self.add_to_list_widget()
        except UnboundLocalError:
            pass

    def filter_shorcuts(self):
        global file_names
        for i in file_names:
                if self.file_name_processor(i,'get_ext') =='.lnk':
                    file_names.remove(i)
        self.validate_file_extensions()

    def validate_file_extensions(self):
        global convert_only
        global merge
        global file_names
        try:
            if convert_only:
                for i in file_names:
                    if self.file_name_processor(i,'get_ext') in ['.jpg','.png','.jpeg']:
                        continue

                    else:
                        while self.file_name_processor(i,'get_ext') not in ['.jpg','.png','.jpeg']:
                            file_names.remove(i)
            else:
                for i in file_names:
                    if self.file_name_processor(i,'get_ext') in ['.jpg','.png','.jpeg','.pdf']:
                        continue
                    else:
                        while self.file_name_processor(i,'get_ext') not in ['.jpg','.png','.jpeg','.pdf']:
                            file_names.remove(i)
            if merge:
                 for i in file_names:
                    if self.file_name_processor(i,'get_ext') in ['.jpg','.png','.jpeg','.pdf']:
                        continue
                    else:
                        while self.file_name_processor(i,'get_ext') not in ['.jpg','.png','.jpeg','.pdf']:
                            file_names.remove(i)
            self.update_banner("[i] Seraching ......... \nDone..... Found {} valid file(s)".format(len(file_names)))
        except ValueError:
            print("")
        s = set(file_names)
        file_names = list(s)
        self.add_to_list_widget()

    def add_to_list_widget(self):
        global file_names
        self.listWidget.clear()
        for i in set(file_names):
            self.listWidget.addItem(i)    

    def remove_list_item(self, list_entity):
        self.listWidget.takeItem(self.listWidget.row(list_entity))
        print(list_entity)
        global file_names
        print(list_entity.text())
        file_names.remove(list_entity.text())
        print(file_names)

#Copy-cat function can be swapped with self.file_name_processor()
    def get_file_name(self,file):
        base=os.path.basename(file)
        name=os.path.splitext(base)[0] 
        return name

    def update_banner(self,ch):
        self.plainTextEdit.setPlainText(ch)

    def convert_to_pdf(self,file_path, do_merge = False):
        img_path = file_path
        global output_path
        pdf_out_path = output_path  + self.get_file_name(file_path) + ".pdf"
        merge_out_path = output_path +"temp/" + self.get_file_name(file_path) + ".pdf"
        print(merge_out_path)
        # Create files if they dont exist
        if (not os.path.exists(pdf_out_path)) and (do_merge == False):
            open(pdf_out_path, 'w').close()
        if (not os.path.exists(merge_out_path)) and (do_merge == True):
            open(merge_out_path, 'w').close()
        image = Image.open(img_path)
        # Deal with alpha channel 
        if image.mode in ('RGBA','LA') or (image.mode =='P' and 'transparency' in image.info):
            if do_merge == False:
                self.process_alpha_channel(image,pdf_out_path)
            else:
                self.process_alpha_channel(image,merge_out_path)
        else: 
            pdf_bytes = img2pdf.convert(image.filename, strict = False) 
            if do_merge:
                file = open(merge_out_path, "wb+")
            else:
                file = open(pdf_out_path, "wb+") 
            file.write(pdf_bytes) 
            file.close()
            image.close() 


    def process_alpha_channel(self,image,pdf_out_path):
        print('process_alpha_channel')
       #convert to rgba if LA format is present due to a bug in pil
        alpha = image.convert('RGBA').split()[-1]
        image_temp = Image.new('RGB', image.size,(255,255,255))
        image_temp.paste(image, mask=alpha)
        image_temp.save(pdf_out_path,'PDF', resolution = 100)    

    def process_files(self):
        global progress_precentage
        global convert_only
        global merge
        global file_names
        if convert_only:   #only image files are kept under this condition
            n_jobs =0
            for i in file_names:
               bp = os.path.basename(i)
               ext=os.path.splitext(bp)[-1]
               if ext in ['.jpg','.png','.bmp','.jpeg']: #Deal with images first
                   self.convert_to_pdf(i)
                   n_jobs += 1
                   self.update_progressbar(n_jobs)
        if merge:
            self.merge_files()

        if (not convert_only) and (not merge):
            n_jobs =0

            for i in file_names:
                print('Split')
                print(i)
                if self.file_name_processor(i,'get_ext') in ['.jpg','.png','.bmp','.jpeg']:
                    file_names.remove(i)
                self.split_pdf(i)
                n_jobs += 1
                self.update_progressbar(n_jobs)
        self.state_reset()
        self.update_banner('Done........')
			  
    def update_progressbar(self,n_jobs_done):
        global file_names
        global progress_precentage
        total_jobs = len(file_names)
        incr = 100 / total_jobs
        jobs_remaining = total_jobs - n_jobs_done
        if jobs_remaining == 0:
            progress_precentage = 100
            self.progressBar.setProperty("value", progress_precentage)
            progress_percentage = 0
        else:
            progress_percentage = int(incr * n_jobs_done) * 100
            self.progressBar.setProperty("value", progress_precentage)
            
    def merge_files(self):
        global file_names
        merger = PdfFileMerger() 
        EOF_MARKER = b'%%EOF'
        for filename in file_names:
            if self.file_name_processor(filename,'get_ext') in ['.jpg','.png','.bmp','.jpeg']:
                print("convert")
                self.convert_to_pdf(filename, do_merge = True)
                new_file_path = output_path +"temp/" + self.get_file_name(filename) + ".pdf"
                file_names = [new_file_path if x==filename else x for x in file_names]
                print(file_names)

        for filename in file_names:
            with open(filename, 'rb') as f:
                contents = f.read()
        # check if EOF is somewhere else in the file
                if EOF_MARKER in contents:
    # we can remove the early %%EOF and put it at the end of the file
                    contents = contents.replace(EOF_MARKER, b'')
                    contents = contents + EOF_MARKER
                else:
    # Some files really don't have an EOF marker
    # In this case it helped to manually review the end of the file
                    print(contents[-8:]) # see last characters at the end of the file
    # printed b'\n%%EO%E'
                    contents = contents[:-6] + EOF_MARKER
                    with open(filename, 'wb') as f:
                        f.write(contents)
            merger.append(PdfFileReader(open(filename, 'rb')))
            final_path = output_path + self.file_name_processor( file_names[0]) + "_merged.pdf"
        merger.write(final_path)
        self.progressBar.setProperty("value", 100)
        file_names.clear()

    def split_pdf(self,pdf_file):
        global output_path
        inputpdf = PdfFileReader(open(pdf_file, "rb"),strict=False)
        fname = self.file_name_processor(pdf_file,'get_name')
        for i in range(inputpdf.numPages):
            output = PdfFileWriter()
            output.addPage(inputpdf.getPage(i))
            if not os.path.exists(output_path + "/" +fname):
                os.mkdir(output_path + "/" +fname)
            print(output_path, fname)
            with open("{}/{}/{}-page{}.pdf".format(output_path, fname, fname, i), "wb") as outputStream:
                output.write(outputStream)

#------------------------------------------------File Explorer Section/ functions to try--------------------------------------------
'''def openFileNameDialog(self):
    options = QtWidgets.QFileDialog.Options()
    options |= QtWidgets.QFileDialog.DontUseNativeDialog
    fileName, _ = QtWidgets.QFileDialog.getOpenFileName(None,"QFileDialog.getOpenFileName()", "","All Files (*);;Python Files (*.py)", options=options)
    if fileName:
        print(fileName)
    
def openFileNamesDialog(self):
    options = QtWidgets.QFileDialog.Options()
    options |= QtWidgets.QFileDialog.DontUseNativeDialog
    files, _ = QtWidgets.QFileDialog.getOpenFileNames(None,"QFileDialog.getOpenFileNames()", "","All Files (*);;Python Files (*.py)", options=options)
    if files:
        print(files)
    
def saveFileDialog(self):
    options = QtWidgets.QFileDialog.Options()
    options |= QtWidgets.QFileDialog.DontUseNativeDialog
    fileName, _ = QtWidgets.QFileDialog.getSaveFileName(None,"QFileDialog.getSaveFileName()","","All Files (*);;Text Files (*.txt)", options=options)
    if fileName:
        print(fileName)
'''



#--------------------------------------------------Run Section----------------------------------------------------        
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    PDFU = QtWidgets.QMainWindow()
    ui = Ui_PDFU()
    ui.setupUi(PDFU)
    PDFU.show()
    sys.exit(app.exec_())
