import os
import re
import random
import string
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
from win32com import client as win32_client

def extract_info_from_pdf(pdf_file):
    with open(pdf_file, 'rb') as file:
        reader = PdfReader(file)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    return text

def extract_info_from_docx(docx_file):
    doc = Document(docx_file)
    text = ''
    for para in doc.paragraphs:
        text += para.text
    return text

def extract_info_from_pdf_file(pdf_file):
    with open(pdf_file, 'rb') as file:
        reader = PdfReader(file)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    return text

def generate_random_email(name):
    # Generate a random 4-digit number
    random_number = ''.join(random.choices(string.digits, k=4))
    # Construct a Gmail address with the random number
    return f"{name.lower().replace(' ', '')}{random_number}@gmail.com"

def extract_email(text):    
    emails = re.findall(r'(?:\bE-Mailid-)?([\w\.-]+@[\w\.-]+(?:\.com)\b)', text, re.IGNORECASE)    
    cleaned_emails = [re.sub(r'\d$', '', email) for email in emails]
    return list(set(cleaned_emails))


def extract_phone_number(text):
    phone_numbers = re.findall(r'[\+\(]?[1-9]\d{0,2}[\)-]?\s*?\d{2,4}[\s.-]?\d{2,4}[\s.-]?\d{2,4}', text)
    unique_numbers = set()
    for number in phone_numbers:
        formatted_number = re.sub(r'[\s.-]', '', number)
        if len(formatted_number) == 10:
            unique_numbers.add(formatted_number)
        elif formatted_number.startswith('+') and len(formatted_number) > 10:
            unique_numbers.add(formatted_number)
    return ', '.join(unique_numbers)

def convert_doc_to_pdf(doc_file):
    # Convert .doc to .pdf
    pdf_file = f"{os.path.splitext(doc_file)[0]}.pdf"
    word = win32_client.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_file)
    doc.SaveAs(pdf_file, FileFormat=17)  # 17 for .pdf format
    doc.Close()
    word.Quit()
    return pdf_file


def process_cv(cv_folder):
    data = []
    for filename in os.listdir(cv_folder):
        file_path = os.path.join(cv_folder, filename)
        try:
            if filename.endswith('.pdf'):
                text = extract_info_from_pdf(file_path)
            elif filename.endswith('.docx'):
                text = extract_info_from_docx(file_path)
            elif filename.endswith('.doc'):
                # Convert .doc to .pdf
                pdf_file = convert_doc_to_pdf(file_path)
                # Extract text from the converted .pdf file
                text = extract_info_from_pdf_file(pdf_file)
                # Remove the temporary .pdf file
                os.remove(pdf_file)
            else:
                continue
            email = extract_email(text)
            if not email:  # If no email is extracted
                # Generate a random email address
                random_email = generate_random_email(filename.split('.')[0])
                email.append(random_email)
            phone_number = extract_phone_number(text)
            # Remove file extension from the filename
            name = os.path.splitext(filename)[0]
            data.append({'Name': name, 'Email': email, 'Phone Number': phone_number, 'Text': text})
        except Exception as e:
            print(f"Error processing {filename}: {e}")
            continue
    return data

def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df['Email'] = df['Email'].apply(lambda x: ', '.join(x))
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    cv_folder = "C:\\Users\\Nandini\\OneDrive\\Desktop\\cv's\\Sample2"  # Full path to the folder
    output_file = "output.xlsx"
    cv_data = process_cv(cv_folder)
    save_to_excel(cv_data, output_file)
