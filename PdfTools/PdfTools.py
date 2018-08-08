import MainForm

from System.Drawing import *
from System.Windows.Forms import *

Application.EnableVisualStyles()
Application.SetCompatibleTextRenderingDefault(False)

form = MainForm.MyForm()
Application.Run(form)
