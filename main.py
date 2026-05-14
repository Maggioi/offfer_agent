import docx

doc = docx.Document("krolik_doswiadczalny.docx")
#(len(doc.paragraphs))

print(len(doc.tables))

for row in doc.tables[2].rows:
    for cell in row.cells:
        print(cell.text)