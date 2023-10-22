from docx import Document
document = Document("new_test.docx")


heading1 = ["chair letters", "history of the committee", "chair letter", 
	"crisis director letter", "committee structure and mechanics", "topic of committee", 
	"powers of the committee", "sensitivity statement", "topic a","topic b"] 

heading2 = ["statement of the problem", "history of the problem",
	"past actions", "bloc positions", "possible solutions", "subhead1", 
	"character bios"]
# individually code subhead

heading2g = ["glossary"]
heading2b = ["bibliography"] #set up flag for bib text

heading3 = ["subhead2"]
letters = ["chair letters", "crisis director letter", "chair letter"]

def delete_paragraph(paragraph):
	p = paragraph._element
	p.getparent().remove(p)
	p._p = p._element = None

for paragraph in document.paragraphs:
	#print(text)
	text = paragraph.text.strip().lower()

	if len(text) == 0:
		delete_paragraph(paragraph)
	
	paragraph.style = "MUNUCNormal"

	
	bib_check = True
	letter_check = True

	for heading in heading1:
		if heading in text:
			paragraph.text = text.replace("topic of committee", "").title() #check if it is title???
			paragraph.style = "Heading 1"
			letter = False
			if text in letters:
				letter = True
				letter_check = False
			bib = False
	if letter and letter_check:
		paragraph.style = "Letters"

	if "subhead0" in text:
		text = text.replace("subhead0", "")
		if "cap1" not in text:
			paragraph.text = text.title()
		else:
			text = text.replace("cap1", "")
			paragraph.text = text.upper()
		paragraph.style = "Heading 1"

	for heading in heading2:
		if heading == text:
			if "cap1" not in text:
				paragraph.text = text.title()
			else:
				text = text.replace("cap1", "")
				paragraph.text = text.upper()
			paragraph.style = "Heading 2"

	if "subhead1" in text:
		text = text.replace("subhead1", "")
		if "cap1" not in text:
			paragraph.text = text.title()
		else:
			text = text.replace("cap1", "")
			paragraph.text = text.upper()
		paragraph.style = "Heading 2"

	if "glossary" == text:
		paragraph.text = "Glossary"
		paragraph.style = "Heading 2G"

	if "bibliography" == text:
		paragraph.text = "Bibliography"
		paragraph.style = "Heading 2B"
		bib = True
		bib_check = False

	if bib and bib_check:
		paragraph.style = "Bibliography"

	if "subhead2" in text:
		text = text.replace("subhead2", "")
		if "cap1" not in text:
			paragraph.text = text.title()
		else:
			text = text.replace("cap1", "")
			paragraph.text = text.upper()
		paragraph.style = "Heading 3"




document.save("test_formatted.docx")


