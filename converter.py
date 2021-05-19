import os.path
import win32com.client
import glob

baseDir = 'C:\\Users\\KyleRid\\Desktop\\py' # Starting directory for directory walk

word = win32com.client.Dispatch("Word.application")
# doc = word.Documents.Open('C:\\Users\\KyleRid\\Desktop\\py\\doc1.doc')
# doc.SaveAs('C:\\Users\\KyleRid\\Desktop\\py\\doc1.docx', 16)
# print(glob.glob(baseDir+'/source_test/doc/*.doc'))
docFiles = glob.glob(baseDir+'/source/doc/*.doc')

for i in docFiles:
    doc = word.Documents.Open(i)
    doc.SaveAs(i.replace('source/doc', 'source/docx')+'x', 16)
    word.ActiveDocument.Close()
    print(i)
# for dir_path, dirs, files in os.walk(baseDir):
#     for file_name in files:

#         file_path = os.path.join(dir_path, file_name)
#         file_name, file_extension = os.path.splitext(file_path)

#         if "~$" not in file_name:
#             if file_extension.lower() == '.doc': #
#                 docx_file = '{0}{1}'.format(file_path, 'x')

#             if not os.path.isfile(docx_file): # Skip conversion where docx file already exists

#                 file_path = os.path.abspath(file_path)
#                 docx_file = os.path.abspath(docx_file)
#                 try:
#                     wordDoc = word.Documents.Open(file_path)
#                     wordDoc.SaveAs2(docx_file, FileFormat = 16)
#                     wordDoc.Close()
#                 except Exception as e:
#                     print('Failed to Convert: {0}'.format(file_path))
#                     print(e)
