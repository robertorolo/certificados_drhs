#!pip install python-docx
#start from here if python-docx is installed
from docx import Document
#open the document
doc=Document('./test.docx')
Dictionary = {"sea": "ocean", "find_this_text":"new_text"}
for i in Dictionary:
    for p in doc.paragraphs:
        if p.text.find(i)>=0:
            p.text=p.text.replace(i,Dictionary[i])
#save changed document
doc.save('./test.docx')
