import os
import fitz
import camelot
import pandas as pd

def search_word_in_pdf(pdf_path, search_word):
    found_pages = []
    document = fitz.open(pdf_path)
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text = page.get_text()
        if search_word.lower() in text.lower():
            found_pages.append(page_num + 1)  # Page numbers are 1-based
    return found_pages

def extract_tables_from_pdf(pdf_path, pages='1-end', flavor='lattice'):
    tables = camelot.read_pdf(pdf_path, pages=pages, flavor=flavor, 
                              line_scale=40, shift_text=[''], strip_text='\n', split_text=True)
    return tables

def save_tables_to_excel(tables, excel_path):
    with pd.ExcelWriter(excel_path) as writer:
        for i, table in enumerate(tables):
            table.df.to_excel(writer, sheet_name=f'Table_{i}', index=False)

def process_all_pdfs_in_folder(folder_path, search_word):
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        excel_file = os.path.join(folder_path, pdf_file.replace('.pdf', '_tables.xlsx'))
        
        # Step 1: Search for the word in the PDF
        found_pages = search_word_in_pdf(pdf_path, search_word)

        if found_pages:
            print(f'The word "{search_word}" was found in "{pdf_file}" on the following pages: {found_pages}')

            # Step 2: Extract tables from specific pages where the word is found
            pages_to_extract = ','.join(map(str, found_pages))
            tables = extract_tables_from_pdf(pdf_path, pages=pages_to_extract, flavor='lattice')

            if tables:
                # Step 3: Save extracted tables to Excel
                save_tables_to_excel(tables, excel_file)
                print(f'Tables extracted from "{pdf_file}" and saved to "{excel_file}" successfully.')
            else:
                print(f"No tables found in the specified pages of \"{pdf_file}\".")
        else:
            print(f'The word "{search_word}" was not found in "{pdf_file}".')

if __name__ == "__main__":
    folder_path = "/workspaces/DataExtraction"  # Replace with the path to your folder
    search_word = "case"
    
    process_all_pdfs_in_folder(folder_path, search_word)