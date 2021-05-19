
# import pandas as pd
from docx.api import Document
import glob



baseDir = 'C:\\Users\\KyleRid\\Desktop\\py\\source\\docx' # Starting directory for directory walk

docFiles = glob.glob(baseDir+'/doc*.docx') # for the loop

header = ()
data = []
counter = 0

for i in docFiles:
    counter = 0
    print(i)
    document = Document(i)
    table = document.tables[0]

    keys = None
    for i, row in enumerate(table.rows):
        text = (cell.text for cell in row.cells)
        if i == 0:
            keys = tuple(text)
            continue
        counter = counter + 1
        row_data = list(zip(keys, text))
        if not header and counter == 1:
            header = row_data
        if counter == 2:
            data.append(row_data)
        print(counter)

unique = []

for i in data:
    if i not in unique:
        unique.append(i)
    else:
        print(i,end=' ')

targetDoc = Document('C:\\Users\\KyleRid\\Desktop\\py\\source\\target.docx')
targetTable = targetDoc.tables[0]
globalCounter = 1

# data = unique
data.sort(key = lambda x: x[3]) #in place sorting


print(len(data))

for item in data:
    print(globalCounter)
    cells = targetTable.add_row().cells
    cells[0].text = str(globalCounter)
    cells[1].text = item[1][1]
    cells[2].text = item[2][1]
    cells[3].text = item[3][1]
    cells[4].text = item[4][1]
    cells[5].text = item[5][1]
    cells[6].text = item[6][1]
    cells[7].text = item[7][1]
    cells[8].text = item[8][1]
    cells[9].text = item[9][1]
    globalCounter = globalCounter + 1

targetDoc.save('C:\\Users\\KyleRid\\Desktop\\py\\source\\target2.docx')











#copy
# header = ()
# data = []
# counter = 0
# keys = None
# for i, row in enumerate(table.rows):
#     text = (cell.text for cell in row.cells)
#     if i == 0:
#         keys = tuple(text)
#         continue
#     counter = counter + 1
#     row_data = dict(zip(keys, text))
#     if not header and counter == 1:
#         header = row_data
#     if counter == 2:
#         data.append(row_data)

# print(data)










# print (header)

# df = pd.DataFrame(data)
