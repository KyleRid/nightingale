import os.path
import win32com.client
import glob

baseDir = 'C:\\Users\\KyleRid\\Desktop\\py' # Starting directory for directory walk

word = win32com.client.Dispatch("Word.application")
docFiles = glob.glob(baseDir+'/source/doc/*.doc')

for i in docFiles:
    doc = word.Documents.Open(i)
    doc.SaveAs(i.replace('source/doc', 'source/docx')+'x', 16)
    word.ActiveDocument.Close()
    print(i)
