import clr
import os
import re
import PdfProvider

clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from System.Drawing import *
from System.Windows.Forms import *
from System.ComponentModel import BackgroundWorker
from PdfProvider import PdfTools

class MyForm(Form):
    def __init__(self):
        self.Text = "PDF Tools"
        self.Icon = Icon("icon.ico")
        self.Height = 450
        self.Width = 700
        self.MaximizeBox = False
        self.FormBorderStyle = FormBorderStyle.FixedSingle
        self.StartPosition = FormStartPosition.CenterScreen
        self.Load += self.formLoad

        self.destinationFile = ""
        self.pageRange = []

        self.backgroundWorker = BackgroundWorker()
        self.backgroundWorker.WorkerReportsProgress = True
        self.backgroundWorker.WorkerSupportsCancellation = True
        self.backgroundWorker.DoWork += self.backgroundWorkerDoWork
        self.backgroundWorker.ProgressChanged += self.backgroundWorkerProgressChanged
        self.backgroundWorker.RunWorkerCompleted += self.backgroundWorkerRunWorkerComplete
        
        self.rbMerge = RadioButton()
        self.rbMerge.Location = Point(15, 320)
        self.rbMerge.Size = Size(300,17)
        self.rbMerge.Text = "Merge files by creating a new file or overwrite an existing file."
        self.rbMerge.Checked = True
        
        self.rbAppend = RadioButton()
        self.rbAppend.Location = Point(15, 340)
        self.rbAppend.Size = Size(300,17)
        self.rbAppend.Text = "Append files to an existing file"
        self.rbAppend.Checked = False
        self.rbAppend.CheckedChanged += self.rbAppendCheckedChanged

        self.progressBar = ProgressBar()
        self.progressBar.Location = Point(405, 332)
        self.progressBar.Size = Size(190, 23)
        self.progressBar.Step = 1
        
        self.lblStatus = Label()
        self.lblStatus.Location = Point(15, 5)
        self.lblStatus.Size = Size(650, 23)
        self.lblStatus.Text = "Destination File: Not set yet!"

        self.lblPercent = Label()
        self.lblPercent.Location = Point(600, 336)
        self.lblPercent.Size = Size(50, 23)
        self.lblPercent.Text = "0%"

        self.lblPageNumber = Label()
        self.lblPageNumber.Enabled = False
        self.lblPageNumber.Location = Point(30, 360)
        self.lblPageNumber.Size = Size(292, 13)
        self.lblPageNumber.Text = "After a specific page number (Zero means at the beginning): "
        
        self.image = Image.FromFile("info.ico")

        self.lblInfo = Label()
        self.lblInfo.Enabled = False
        self.lblInfo.Location = Point(400, 388)
        self.lblInfo.Size = Size(50, 17)
        self.lblInfo.Image = self.image

        self.rbSplit = RadioButton()
        self.rbSplit.Location = Point(15, 380)
        self.rbSplit.Size = Size(300,17)
        self.rbSplit.Text = "Split page(s) and save in new file: "
        self.rbSplit.Checked = False
        self.rbSplit.CheckedChanged += self.rbSplitCheckedChanged
        
        self.tbPagesRange = TextBox()
        self.tbPagesRange.Enabled = False
        self.tbPagesRange.Location = Point(322, 384)
        self.tbPagesRange.MaxLength = 12
        self.tbPagesRange.Size = Size(80, 20)
        self.tbPagesRange.KeyPress += self.tbPagesRangeOnKeyPress
        
        self.textBoxPageNum = TextBox()
        self.textBoxPageNum.Enabled = False
        self.textBoxPageNum.Location = Point(322, 357)
        self.textBoxPageNum.MaxLength = 3
        self.textBoxPageNum.Size = Size(80, 20)
        self.textBoxPageNum.Text = ""
        self.textBoxPageNum.TextAlign = HorizontalAlignment.Center
        self.textBoxPageNum.KeyPress += self.textBoxPageNumOnKeyPress

        self.btnSelectFiles = Button()
        self.btnSelectFiles.Location = Point(12, 22)
        self.btnSelectFiles.Size = Size(109, 23)
        self.btnSelectFiles.Text = "Select PDF Files"
        self.btnSelectFiles.UseVisualStyleBackColor = True
        self.btnSelectFiles.Click += self.btnSelectFilesClick

        self.btnDestination = Button()
        self.btnDestination.Location = Point(127, 22)
        self.btnDestination.Size = Size(109, 23)
        self.btnDestination.Text = "Destination File"
        self.btnDestination.UseVisualStyleBackColor = True
        self.btnDestination.Click += self.btnDestinationClick
        
        self.btnUp = Button()
        self.btnUp.Location = Point(600, 50)
        self.btnUp.Size = Size(75, 23)
        self.btnUp.Text = "Up"
        self.btnUp.UseVisualStyleBackColor = True
        self.btnUp.Click += self.btnUpClick
        
        self.btnDown = Button()
        self.btnDown.Location = Point(600, 80)
        self.btnDown.Size = Size(75, 23)
        self.btnDown.Text = "Down"
        self.btnDown.UseVisualStyleBackColor = True
        self.btnDown.Click += self.btnDownClick
  
        self.btnRun = Button()
        self.btnRun.Location = Point(520, 358)
        self.btnRun.Size = Size(75, 23)
        self.btnRun.Text = "Run"
        self.btnRun.UseVisualStyleBackColor = True
        self.btnRun.Click += self.btnRunClick

        self.btnDelete = Button()
        self.btnDelete.Location = Point(436, 22)
        self.btnDelete.Size = Size(75, 23)
        self.btnDelete.Text = "Delete"
        self.btnDelete.UseVisualStyleBackColor = True
        self.btnDelete.Click += self.btnDeleteClick

        self.btnDeleteAll = Button()
        self.btnDeleteAll.Location = Point(518, 22)
        self.btnDeleteAll.Size = Size(75, 23)
        self.btnDeleteAll.Text = "Delete All"
        self.btnDeleteAll.UseVisualStyleBackColor = True
        self.btnDeleteAll.Click += self.btnDeleteAllClick
        
        self.listBoxFiles = ListBox()
        self.listBoxFiles.FormattingEnabled = True
        self.listBoxFiles.Location = Point(12, 51)
        self.listBoxFiles.Size = Size(582, 265)

        self.Controls.Add(self.listBoxFiles)
        self.Controls.Add(self.btnRun)
        self.Controls.Add(self.btnDown)
        self.Controls.Add(self.btnUp)
        self.Controls.Add(self.btnSelectFiles)
        self.Controls.Add(self.btnDelete)
        self.Controls.Add(self.btnDeleteAll)
        self.Controls.Add(self.btnDestination)
        self.Controls.Add(self.progressBar)
        self.Controls.Add(self.rbMerge)
        self.Controls.Add(self.rbAppend)
        self.Controls.Add(self.lblStatus)
        self.Controls.Add(self.lblPercent)
        self.Controls.Add(self.lblPageNumber)
        self.Controls.Add(self.textBoxPageNum)
        self.Controls.Add(self.rbSplit)
        self.Controls.Add(self.tbPagesRange)
        self.Controls.Add(self.lblInfo)

    # # # # # # # EVENTS # # # # # # #
    
    def formLoad(self, sender, e):
        self.toolTip = ToolTip()
        self.toolTip.AutoPopDelay = 10000
        self.toolTip.SetToolTip(self.lblInfo, "Page Range: 3 or 3,5,7 or 3-7")

    # Add PDF files to ListBox
    def btnSelectFilesClick(self, sender, e):
        openFileDialog = OpenFileDialog()
        openFileDialog.Filter = "PDF files (*.pdf)|*.pdf"
        openFileDialog.Multiselect = True
        
        if openFileDialog.ShowDialog():
            for fileName in openFileDialog.FileNames:
                if not fileName.lower().endswith(".pdf"):
                    MessageBox.Show("Selected file: " + fileName + " is not PDF file!", "Error")
                elif fileName == self.destinationFile:
                    MessageBox.Show("Selected file: " + fileName + " is already as destination file", "Error")
                else:
                    self.listBoxFiles.Items.Add(fileName)
   
    # Set destinatio file
    def btnDestinationClick(self, sender, e):
        saveFileDialog = SaveFileDialog()
        saveFileDialog.Filter = "PDF files (*.pdf)|*.pdf"
        
        if saveFileDialog.ShowDialog():
            for item in self.listBoxFiles.Items:
                if saveFileDialog.FileName == item:
                    MessageBox.Show(saveFileDialog.FileName + \
                        " is already exist in the selected files list, try another file name", \
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    return
        
            if saveFileDialog.FileName.lower().endswith('.pdf') == False:
                MessageBox.Show(self.destinationFile + \
                    " extension file doesn't has PDF extension, try again!", \
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                return
            
            self.destinationFile = saveFileDialog.FileName
            self.lblStatus.Text = "Destination File: " + self.destinationFile

    # Move selected item in list box up
    def btnUpClick(self, sender, e):
        if self.listBoxFiles.SelectedIndex > 0:
            self.listBoxFiles.Items.Insert(self.listBoxFiles.SelectedIndex - 1, self.listBoxFiles.SelectedItem)
            self.listBoxFiles.SelectedIndex = self.listBoxFiles.SelectedIndex - 2
            self.listBoxFiles.Items.RemoveAt(self.listBoxFiles.SelectedIndex + 2)
    
    # Move selected item in list box down
    def btnDownClick(self, sender, e):
        if self.listBoxFiles.SelectedIndex != -1 and self.listBoxFiles.SelectedIndex < self.listBoxFiles.Items.Count - 1:
            self.listBoxFiles.Items.Insert(self.listBoxFiles.SelectedIndex + 2, self.listBoxFiles.SelectedItem)
            self.listBoxFiles.SelectedIndex = self.listBoxFiles.SelectedIndex + 2
            self.listBoxFiles.Items.RemoveAt(self.listBoxFiles.SelectedIndex - 2)
    
    # Remove selected item in list box
    def btnDeleteClick(self, sender, e):
         if self.listBoxFiles.SelectedIndex != -1:
             self.listBoxFiles.Items.RemoveAt(self.listBoxFiles.SelectedIndex)

    # Remove all list box items
    def btnDeleteAllClick(self, sender, e):
         if self.listBoxFiles.Items.Count > 0:
            self.listBoxFiles.Items.Clear()

    # Set some prerequisites and run the backgroud worker!
    def btnRunClick(self, sender, e):
        self.progressBar.Value = 0
        self.progressBar.Maximum = self.listBoxFiles.Items.Count
        
        if self.setParameters():
            self.disableButtons()
            self.progressBar.Value = 0
            if self.pageRange.Count > 0:
                self.progressBar.Maximum = self.pageRange.Count
            else:
                self.progressBar.Maximum = self.listBoxFiles.Items.Count
            
            if self.backgroundWorker.IsBusy != True:
                self.backgroundWorker.RunWorkerAsync()
    
    # Enable and disable split text box
    def rbSplitCheckedChanged(self, sender, e):
        if self.rbSplit.Checked:
            self.tbPagesRange.Enabled = True
            self.lblInfo.Enabled = True
        else:
            self.tbPagesRange.Enabled = False
            self.lblInfo.Enabled = False
    
    # Enable and disable append text box
    def rbAppendCheckedChanged(self, sender, e):
        if self.rbAppend.Checked:
            self.lblPageNumber.Enabled = True
            self.textBoxPageNum.Enabled = True
        else:
            self.lblPageNumber.Enabled = False
            self.textBoxPageNum.Enabled = False
    
    # Prevent to enter non numerical and white space chars
    def textBoxPageNumOnKeyPress(self, sender, e):
        keyChar = str(e.KeyChar)          
        match = re.match(r"[^0-9]", keyChar)

        # Backspace is allowed
        if ord(keyChar) == 8:
            return

        if keyChar.isspace() or match:
            e.Handled = True
    
    # Prevent to enter non numerical and white space chars
    def tbPagesRangeOnKeyPress(self, sender, e):
        keyChar = str(e.KeyChar)          
        match = re.match(r"[^0-9,-]", keyChar)
        
        # Backspace is allowed
        if ord(keyChar) == 8:
            return
            
        if keyChar.isspace() or match:
            e.Handled = True

    # Determine which PDF method in PdfProvider should call.
    def backgroundWorkerDoWork(self, sender, e):
        bgWorker = sender
        
        if self.rbMerge.Checked:
            self.merginJob(self.listBoxFiles.Items, self.destinationFile, bgWorker, e)
        elif self.rbSplit.Checked:
             self.splitJob(self.listBoxFiles.Items[0], self.destinationFile, self.pageRange, bgWorker, e)
        elif self.rbAppend.Checked and len(self.textBoxPageNum.Text) == 0:
            self.appendingJob(self.listBoxFiles.Items, self.destinationFile, bgWorker, e)
        elif len(self.textBoxPageNum.Text) > 0:
            self.appendingJobPage(self.listBoxFiles.Items, self.destinationFile, self.textBoxPageNum.Text, \
                bgWorker, e)

    # Update progress bar and precent lable after each file merged or appended!
    def backgroundWorkerProgressChanged(self, sender, e):
         # Solution founded to fix the progress bar changing value delay!  #
         if e.ProgressPercentage == self.progressBar.Maximum:
             self.progressBar.Maximum = e.ProgressPercentage + 1              
             self.progressBar.Value = e.ProgressPercentage + 1
             self.progressBar.Maximum = e.ProgressPercentage
         else:
            self.progressBar.Value = e.ProgressPercentage + 1
            
         self.progressBar.Value = e.ProgressPercentage
        #########################################################################
         self.lblPercent.Text = str(int(self.progressBar.Value * 100 / self.progressBar.Maximum)) + "%"

    # Background worker error handling, and re enable text boxes.
    def backgroundWorkerRunWorkerComplete(self, sender, e):
        if e.Error != None:
            MessageBox.Show(e.Error.Message)
        self.enableButtons()

    # # # # # # # Helper Methods # # # # # # #

    # Set parameters base on selected job.
    def setParameters(self):
        checkedResult = False

        if self.checkDestination():
            if self.rbMerge.Checked:
                checkedResult = self.checkMergeParameters()
            elif self.rbAppend.Checked:
                checkedResult = self.checkAppendParameters()
            else:
                checkedResult = self.checkSplitParameters()
        
        return checkedResult
    
    # Check selected destination file
    def checkDestination(self):
        if len(self.destinationFile) == 0:
            MessageBox.Show("Destination file is not set!","Error",\
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False      
        return True
    
    # Check validation of mergin parameters 
    def checkMergeParameters(self):
        if self.listBoxFiles.Items.Count < 2:
            MessageBox.Show("Select at least two PDF files to merge them!","Error",\
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False
        
        return True

    # Check validation of appending parameters  
    def checkAppendParameters(self):
        if self.listBoxFiles.Items.Count == 0:
                MessageBox.Show("Select at least one PDF file to append it!","Error",\
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                return False

            # For appeding, the destination file should be exist.
        if not os.path.exists(self.destinationFile):
            MessageBox.Show("Destination file is not exist to append file(s) to it!","Error",\
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False

        pdfTools = PdfTools()
        pdfTools.destinationFile = self.destinationFile

        if len(self.textBoxPageNum.Text) > 0:
            if not pdfTools.checkPageAppending(int(self.textBoxPageNum.Text)):
                MessageBox.Show("Page number is not current!","Error", \
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                return False
                
        return True

      # Check validation of splitting parameters
    def checkSplitParameters(self):
        if self.listBoxFiles.Items.Count == 0:
            MessageBox.Show("Select a PDF file to split pages from it!","Error",\
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False
        
        if self.listBoxFiles.Items.Count > 1:
            MessageBox.Show("Select just one PDF file to split pages from it!","Error",\
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False

        if not os.path.exists(self.listBoxFiles.Items[0]):
            MessageBox.Show("Source file is not exist to split pages from it!","Error",\
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False

        if len(self.tbPagesRange.Text) == 0:
            MessageBox.Show("Page range is empty!","Error",\
                MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False

        patterns = [",,","--",",-","-,","^,","^-",",$","-$","[0-9]+-[0-9]+-",]
        for pattern in patterns:
            if re.search(pattern, self.tbPagesRange.Text):
                MessageBox.Show("Page range is not valid, check out the info icon!","Error",\
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
                return False

        pattern = self.tbPagesRange.Text
        self.pageRange.Clear()

        for x in pattern.split(","):
            if not str(x).Contains("-"):
                self.pageRange.append(int(x))
            else:
                # Like 3-5 or 9-6
                rangeTemp = str(x).split("-")
                n1 = int(rangeTemp[0])
                n2 = int(rangeTemp[1])

                # Extract pages number between the range
                if n1 < n2:
                    for x in range(n1, n2 + 1):
                        self.pageRange.append(int(x))
                elif n1 > n2:
                    for x in range(n1, n2 - 1, -1):
                        self.pageRange.append(int(x))
                else:
                    self.pageRange.append(int(n1))
        
        pdfTools = PdfTools()
        pdfTools.sourceFile = self.listBoxFiles.Items[0]

        if not pdfTools.checkPageSplit(max(self.pageRange), min(self.pageRange)):
            MessageBox.Show("Page number is not current!","Error", \
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
            return False

        return True

    # All buttons should be disable during merging or appending files.
    def disableButtons(self):
        self.btnDelete.Enabled = False
        self.btnDeleteAll.Enabled = False
        self.btnDown.Enabled = False
        self.btnRun.Enabled = False
        self.btnDestination.Enabled = False
        self.btnSelectFiles.Enabled = False
        self.btnUp.Enabled = False

    # Re enable all buttons after mergin or appending files
    def enableButtons(self):
        self.btnDelete.Enabled = True
        self.btnDeleteAll.Enabled = True
        self.btnDown.Enabled = True
        self.btnRun.Enabled = True
        self.btnDestination.Enabled = True
        self.btnSelectFiles.Enabled = True
        self.btnUp.Enabled = True

     # # Mergin and appending job methods, will called by background worker # #

    def merginJob(self, filesList, destinationFilePath, worker, event):
        pdfProvider = PdfTools()
        pdfProvider.destinationFile = destinationFilePath
        counter = 0
        result = ""

        for item in filesList:
            result = pdfProvider.merging(item)

            if not result == "true":
                MessageBox.Show(result, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                return

            counter += 1
            worker.ReportProgress(counter)

    def appendingJob(self, filesList, destinationFilePath, worker, event):
        pdfProvider = PdfTools()
        pdfProvider.destinationFile = destinationFilePath
        counter = 0
        result = ""

        for item in filesList:
            result = pdfProvider.appending(item)

            if not result == "true":
                MessageBox.Show(result, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                return

            counter += 1
            worker.ReportProgress(counter)

    def appendingJobPage(self, filesList, destinationFilePath, pageNumber, worker, event):
         pdfProvider = PdfTools()
         pdfProvider.destinationFile = destinationFilePath
         counter = 0
         result = ""

         for file in filesList:
            result = pdfProvider.appendingAfterPage(file, int(pageNumber))                    
            
            if not result == "true":
                MessageBox.Show(result, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                return

            counter += 1
            worker.ReportProgress(counter)  

    def splitJob(self, sourceFile, destinationFilePath, pagesRange, worker, event):
         pdfProvider = PdfTools()
         pdfProvider.destinationFile = destinationFilePath
         pdfProvider.sourceFile = sourceFile
         result = ""
         counter = 0

         for pageNum in pagesRange:
             result = pdfProvider.splitPages(pageNum)

             if not result == "true":
                 MessageBox.Show(result, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                 return

             counter += 1
             worker.ReportProgress(counter)