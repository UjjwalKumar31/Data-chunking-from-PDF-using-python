import pdfplumber
import pandas as pd
import openpyxl
import shutil
import logging
import os
import re

if os.path.exists(r'output'):
    shutil.rmtree('output')

def extract_bold_headers(pdf_path):
    """Extract bold headers from a PDF file."""
    
    headers = []
    bold_headers = {}
    page_ref = {}
    for filename in os.listdir(pdf_path):
        if filename.endswith('.pdf'):
            page_ref[filename] = {}
            bold_headers[filename] = []
            with pdfplumber.open(pdf_path+'/'+filename) as pdf:
                print(f'progressing with {filename}')
                for page_number, page in enumerate(pdf.pages, start=1):
                    words = page.extract_words()

                    strr = ''
                    for element in words:
                        flag = False

                        file_path = os.path.join('output', 'pdf element.txt')

                        with open(file_path, 'a') as file:
                            file.write(f"\n {element} \n")

                        if element['height'] > 22:   #'height': 22.517250750000017    
                            flag = True

                        if flag == True:                
                            strr += ''.join(element['text'] + ' ')
                        else:
                            if strr:
                                headers.append(strr.strip().replace('\n', ''))
                                page_ref[filename][strr.strip().replace('\n', '')] = page_number
                                
                            strr = ''
            bold_headers[filename] = headers
            headers = []

    return bold_headers, page_ref

def create_and_write_file(folder_path, file_name, text):
    """Create a file in the specified folder and write text to it."""

    os.makedirs(folder_path, exist_ok=True)

    file_path = os.path.join(folder_path, file_name)

    with open(file_path, 'a', encoding='utf-8', errors='ignore') as file:
        file.write(text)

def chunk_text(pdf_path, headers):

    """Extract & Chunk text from a PDF file."""
    chunk_list = []
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"

    indx = 0
    # print(pdf_path)
    while indx < len(headers)-1:

        i,j = text.index(headers[indx][:21]), text.index(headers[indx+1][:21])
        chunk = text[i : j]
        create_and_write_file('output', f'{pdf_path[11:-4]}.txt', f"\nChunk {indx + 1}:\n{chunk}\n")
        indx += 1
        chunk_list.append(chunk)

    print(indx, len(headers), headers[indx], headers[-1])
    if indx < len(headers):
        i = text.index(headers[-1][:15])
        chunk = text[i : ]
        create_and_write_file('output', f'{pdf_path[11:-4]}.txt', f"\nChunk {indx + 1}:\n{chunk}\n")
        chunk_list.append(chunk)

    return chunk_list

def create_QB(header, chunk_list, page_num):
     
    extracted_data = {
          'Questions' : header,
          'chunk_list': chunk_list,
          'page_num': page_num
     }
    df = pd.DataFrame(extracted_data)

    excel_path = "output/extracted_data.xlsx"

    if os.path.exists(r'output\extracted_data.xlsx'):
        df_existing = pd.read_excel("output/extracted_data.xlsx", sheet_name='Master QB')
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name='Master QB', startrow=len(df_existing) + 1, index=False, header=False)
    else:
        
        df.to_excel(excel_path, sheet_name= 'Master QB', index=False, engine='openpyxl')

    print(" Master QB created successfully!")

def main():
    pdf_path = "Documents/About Boots.pdf"
    folder = 'Documents'
    bold_headers, page_ref = extract_bold_headers(folder)

    for pdf, header in bold_headers.items():
        file_path = os.path.join('output', 'pdf headers.txt')

        with open(file_path, 'a') as file:
            file.write(f"\n {pdf}: \n {header} \n")

        print(f"chunck processing for {pdf}")

        chunk = chunk_text(folder+'/'+pdf, header)

        create_QB(header, chunk_list=chunk, page_num=(list(page_ref[pdf].values())))


if __name__ == "__main__":
    os.makedirs('output', exist_ok=True)
    main()
