import os
import shutil
import zipfile
import pandas as pd
import docx2txt
import PyPDF2 
import spacy

# Load English tokenizer, tagger, parser, NER and word vectors
nlp = spacy.load("en_core_web_sm")

# Set up directory for temporary files
path = 'filename34'

# Function to extract text from PDF
def extract_pdf_text(file_path):
    text = ""
    pdfReader = PyPDF2.PdfFileReader(open(file_path, "rb"))
    for page_num in range(pdfReader.numPages):
        pageObj = pdfReader.getPage(page_num)
        text += pageObj.extractText()
    return text

# Function to extract text from DOCX
def extract_docx_text(file_path):
    return docx2txt.process(file_path)

# Check if it's the initial run or a session refresh
init = True
if os.path.exists(path):
    shutil.rmtree(path)
os.makedirs(path)

# Document file location
docx_file_location = input("Enter Document File Location (PDF, DOCX, TXT): ")

# Process uploaded document
if docx_file_location.endswith(".pdf"):
    text = extract_pdf_text(docx_file_location)
elif docx_file_location.endswith(".docx"):
    text = extract_docx_text(docx_file_location)
else:
    with open(docx_file_location, "r") as f:
        text = f.read()

data1 = pd.DataFrame({'Descriptions': [text], 'Name': [os.path.basename(docx_file_location)]})
print(data1)

# ZipFile location
zip_file_location = input("Enter ZipFile Location: ")
if os.path.exists(zip_file_location):
    with zipfile.ZipFile(zip_file_location, "r") as z:
        z.extractall(path)
    FILENAMES = []
    for root, dirs, files in os.walk(path):
        for file in files:
            FILENAMES.append(os.path.join(root, file))

    data = []
    for file in FILENAMES:
        if file.endswith(".docx"):
            text = extract_docx_text(file)
        elif file.endswith(".pdf"):
            text = extract_pdf_text(file)
        data.append({'Descriptions': text, 'Name': os.path.basename(file)})

    data = pd.DataFrame(data)
    print(data)

    # Similarity calculation
    ref_sent = data1['Descriptions'].iloc[0]
    ref_sent_vec = nlp(ref_sent)
    sims = []
    for doc in data['Descriptions']:
        sim = nlp(doc).similarity(ref_sent_vec)
        sims.append(sim)
    data['%_Similarity'] = sims
    sims_docs_sorted = data.sort_values(by='%_Similarity', ascending=False).reset_index(drop=True)
    print(sims_docs_sorted)

    # Save results to Excel
    excel_file = "Text_Similarity_Results.xlsx"
    sims_docs_sorted.to_excel(excel_file, index=False)
    
    print('Text Similarity Results are ready! Saved to', excel_file)
else:
    print("Invalid zip file path.")
