import docx

doc = docx.Document("krolik_doswiadczalny.docx")
print(len(doc.paragraphs))
print(doc.paragraphs[0].text)