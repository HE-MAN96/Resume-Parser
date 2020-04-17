
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTFigure, LTImage, LTTextLineHorizontal, LTChar, LTLine, \
    LTText
import docx2txt
import io
import os

file_path = r"file path "
#file_path1 = r""

##############################Extract Functions#########################

def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as fh:
# iterate over all pages of PDF document
        for page in PDFPage.get_pages(fh, caching=True, check_extractable=True):
# creating a resoure manager
            resource_manager = PDFResourceManager()
# create a file handle
            fake_file_handle = io.StringIO()
# creating a text converter object
            converter = TextConverter(
                resource_manager,
                fake_file_handle,
                codec='utf-8',
                laparams=LAParams()
            )
# creating a page interpreter
            page_interpreter = PDFPageInterpreter(
                resource_manager,
                converter
            )

# process current page
            page_interpreter.process_page(page)

# extract text
            text = fake_file_handle.getvalue()
            yield text

# close open handles
            converter.close()
            fake_file_handle.close()

def extract_text_from_doc(doc_path):
    print("Doc function called")
    data = docx2txt.process(doc_path)
    text = [line.replace('\t', ' ') for line in data.split('\n') if line]
    return ' \n'.join(text)

####################################################styles fucntions##########
def styles_doc(path):
    from docx import Document
    name=[]
    size=[]
    data = Document(path)
    for p in data.paragraphs:
        name.append(p.style.font.name)
        size.append(p.style.font.size)
    return set(name), set(size)

#def styles_pdf():


##########################################calling above functions and extracting text
text=""
if file_path.lower().endswith(('.pdf')):
    for page in extract_text_from_pdf(file_path):
        text += ' ' + page

elif file_path.lower().endswith(('.doc','.docx')):
    for page in extract_text_from_doc(file_path):
        text += ' ' + page
    style1 = styles_doc(file_path)
    print(style1)


###########################functions to extract data#####################
def extract_phone_numbers(string):
    r = re.compile(r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})')
    phone_numbers = r.findall(string)
    numbers = [re.sub(r'\D', '', number) for number in phone_numbers]
    for i in numbers:
        if len(i)>=10:
            return i

def extract_email_addresses(string):
    r = re.compile(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+")
    return r.findall(string)

#####this fn not used anywhere##
def extract_names(txt):
    for sent in nltk.sent_tokenize(txt):
        tokens = nltk.tokenize.word_tokenize(sent)
        tags = ner_tagger.tag(tokens)
        for tag in tags:
            if tag[1] == 'PERSON': return tag




import nltk
import re
from nltk.tag import StanfordNERTagger        #######not used
jar = 'stanford-ner.jar'                      ######not used
model = 'english.muc.7class.distsim.prop'      #####not used
ner_tagger = StanfordNERTagger(model, jar, encoding='utf8')    ######not used

from nltk import word_tokenize
document = "".join(i for i in text.split())
sentences1= nltk.sent_tokenize(document)                        ########not used
sentences2 = [nltk.word_tokenize(sent) for sent in sentences1]  ########not used
sentences = [nltk.pos_tag(sent) for sent in sentences2]          ########not used




###################################################Value extraction############################
phone = extract_phone_numbers(document)
print(phone)
email = extract_email_addresses(document)
print(email)
#name= extract_names(document)
#print(name)
total_characters = len(word_tokenize(text))
print(total_characters)
nlines = text.count('\n')
print(nlines)

