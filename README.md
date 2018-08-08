# simple-pdf-toolkit
Merging, appending and splitting PDF pages.

## Used Items

1. Windows Form Application.
2. Python programming language.
1. Python development in Visual Studio 2017 for debugging Python codes.
2. [IronPython](http://ironpython.net/), the Python programming language for the .NET Framework.
3. [PyPDF2](https://pythonhosted.org/PyPDF2/), pure-python PDF toolkit originating from the pyPdf project.

## Run On Windows

1. Simply open the project in Visual Studio and start it like other projects.
2. In the CMD `ipy C:\address of the project\PdfTools\PdfTools.py`
3. Create a console application and run the project from it like below
``` Command Prompt
   var process = new Process();
   var startInfo =
     new ProcessStartInfo
      {
        WindowStyle = ProcessWindowStyle.Hidden,
        FileName = "CMD.exe",
        Arguments = "/C ipy C:\\project address\\PdfTools\\PdfTools.py"
      };
    process.StartInfo = startInfo;
    process.Start();
``` 
## Notice
If you like to run the project from a console application, the project icons should be exist in the debug folder.
