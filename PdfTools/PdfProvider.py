import os
import cStringIO
import PyPDF2

from PyPDF2 import *

class PdfTools(object):
    def __init__(self):
        self.pdfWriter = PyPDF2.PdfFileWriter()
        self.pdfBeginFile = PyPDF2.PdfFileWriter()
        self.pdfRestFile = PyPDF2.PdfFileWriter()
        self.destinationFile = ""
        self.sourceFile = ""

        # For appending the origin file should read just one time.
        self.isFirstTime = True
    def checkPageAppending(self, maxPageNum):
        try:
            sourceFile = open(self.destinationFile, "rb")
            reader = PyPDF2.PdfFileReader(sourceFile)

            if reader.getNumPages() < maxPageNum:
                sourceFile.close()
                return False
            else:
                sourceFile.close()
                return True
        except :
            with open("log.txt", 'a') as file:
                file.write(e.message + "\n")

    def checkPageSplit(self, maxPageNum, minPageNum):
        try:
            sourceFile = open(self.sourceFile, "rb")
            reader = PyPDF2.PdfFileReader(sourceFile)

            if reader.getNumPages() < maxPageNum or minPageNum == 0:
                sourceFile.close()
                return False
            else:
                sourceFile.close()
                return True
        except :
             with open("log.txt", 'a') as file:
                file.write(e.message + "\n")

    # Merge selected files together.
    def merging(self, file):
        try:
            pdfFile = open(file, "rb")
            pdfReader = PyPDF2.PdfFileReader(pdfFile)
            
            for pageNumber in range(pdfReader.getNumPages()):
                self.pdfWriter.addPage(pdfReader.getPage(pageNumber))

            outputStream = open(self.destinationFile, "wb")
            self.pdfWriter.write(outputStream)

            outputStream.close()
            pdfFile.close()
            return "true"
        except BaseException, e:
            with open("log.txt", 'a') as file:
                file.write(e.message + "\n")
            return str(e)
    
    # Append selected file to an existin file
    def appending(self, pdfFile):   
        try:
            # Buffer in memory to keep the file after write
            memory = cStringIO.StringIO()
            originFile = open(self.destinationFile, "rb") 

            # Read the origin file pages and add them to the pdfWriter buffer
            # in memory just one time.
            if self.isFirstTime and os.path.exists(self.destinationFile):
                originFileReader = PyPDF2.PdfFileReader(originFile)
                for pageNumber in range(originFileReader.getNumPages()):
                    self.pdfWriter.addPage(originFileReader.getPage(pageNumber))

            # Read the current selected file's pages in the listbox and add
            # them to the pdfWriter buffer in memory.  The previous files pages
            # will keep in the buffer until background worker is running.
            selectedFile = open(pdfFile, "rb")
            selectedFileReader = PyPDF2.PdfFileReader(selectedFile)
            for pageNumber in range(selectedFileReader.getNumPages()):
                self.pdfWriter.addPage(selectedFileReader.getPage(pageNumber))
            
            # Write origin file and other files in list to memory buffer as
            # bytes.
            self.pdfWriter.write(memory)
            
            # Remove the origin file, so the memory buffer could save in disk
            # with same path and file name of origin file.
            if os.path.exists(self.destinationFile):
                originFile.close()
                os.remove(self.destinationFile)

            # Save memory buffer in disk with same path and file name of the
            # origin file that removed above.
            outputStream = open(self.destinationFile, "wb")
            outputStream.write(memory.getvalue())
            outputStream.close()

            memory.close()
            selectedFile.close()

            # Set false, so in the next steps, the PDF reader won't read the
            # origin file again
            self.isFirstTime = False 
            return "true"
        except BaseException, e:
            with open("log.txt", 'a') as file:
                file.write(e.message + "\n") 
            return str(e)
    
    # Append seleceted files after a spedific page number in origin file
    def appendingAfterPage(self, pdfFile, pageNo=0):
        try:
            # Memory buffer for the beginning of the destiation file until the
            # specific page number
            memoryBegin = cStringIO.StringIO()

            # Memory buffer for selected files by user to append after the
            # specific file.
            memoryFiles = cStringIO.StringIO()

            # Buffer for the rest of destination file.  from the specific page
            # number to the end of file
            memoryRest = cStringIO.StringIO()

            # Just one time read from beginning of the destination file.
            # From begin to spedific page number and from that page to end of
            # the file, separately
            originFile = open(self.destinationFile, "rb")
            if self.isFirstTime:
                originFileReader = PyPDF2.PdfFileReader(originFile)
                for pageNumber in range(pageNo):
                    self.pdfBeginFile.addPage(originFileReader.getPage(pageNumber))

                for pageNumber in range(pageNo, originFileReader.getNumPages()):
                    self.pdfRestFile.addPage(originFileReader.getPage(pageNumber))

            # Read the pages of the selected file(s) by user
            selectedFile = open(pdfFile, "rb")
            selectedFileReader = PyPDF2.PdfFileReader(selectedFile)
            for pageNumber in range(selectedFileReader.getNumPages()):
                self.pdfWriter.addPage(selectedFileReader.getPage(pageNumber))

            # Write pages to memory buffer
            self.pdfBeginFile.write(memoryBegin)
            self.pdfWriter.write(memoryFiles)
            self.pdfRestFile.write(memoryRest)
            
            # Remove the origin file (destination file) from disk, so
            # the buffer memory could be save in same path and with same name
            # of the origin file.
            originFile.close()   # OriginFile = self.destinationFile -> should be close before remove!
            os.remove(self.destinationFile)

            # Merge the buffers together ...
            merge = PyPDF2.PdfFileMerger()
            merge.append(PyPDF2.PdfFileReader(memoryBegin))
            merge.append(PyPDF2.PdfFileReader(memoryFiles))
            merge.append(PyPDF2.PdfFileReader(memoryRest))

            # Write the merged buffers to disk with path and name of origin
            # file.
            outputStream = open(self.destinationFile, "wb")
            merge.write(outputStream)
            
            # Close everthing!
            memoryBegin.close()
            memoryFiles.close()
            memoryRest.close()
            selectedFile.close()
            outputStream.close()

            # Set false to the method will just read the destiation file just
            # one time.
            self.isFirstTime = False 
            return "true"
        except BaseException, e:
            with open("log.txt", 'a') as file:
                file.write(e.message + "\n")
            return str(e)
            
    # Split selected pages from source file and save them in destination file.
    def splitPages(self, pageNum):
        try:
            # Temp buffer in memory
            memory = cStringIO.StringIO()

            # Read input page number from source file
            originFile = open(self.sourceFile, "rb")
            originFileReader = PyPDF2.PdfFileReader(originFile)
            self.pdfWriter.addPage(originFileReader.getPage(pageNum-1))
            
            # Write page to memory buffer
            self.pdfWriter.write(memory)
            originFile.close()

            # Write buffer to destination file
            outputStream = open(self.destinationFile, "wb")
            outputStream.write(memory.getvalue())
            outputStream.close()

            memory.close()
            return "true"
        except BaseException, e:
            with open("log.txt", 'a') as file:
                file.write(e.message + "\n") 
            return str(e)