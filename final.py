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

def process_all_pdfs_in_folder(folder_path, search_words):
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    pdf_data = {}  # Dictionary to store extracted data for each PDF
    
    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        pdf_data[pdf_file] = {}  # Initialize dictionary for this PDF
        
        for search_word in search_words:
            # Step 1: Search for the word in the PDF
            found_pages = search_word_in_pdf(pdf_path, search_word)

            if found_pages:
                print(f'The word "{search_word}" was found in "{pdf_file}" on the following pages: {found_pages}')
                
                # Step 2: Extract tables from specific pages where the word is found
                for page_num in found_pages:
                    tables = extract_tables_from_pdf(pdf_path, pages=str(page_num), flavor='lattice')
                    if tables:
                        # Step 3: Save extracted tables to Excel
                        excel_file = os.path.join(folder_path, f"{pdf_file.replace('.pdf', '')}_Page_{page_num}_tables.xlsx")
                        save_tables_to_excel(tables, excel_file)
                        print(f'Tables extracted from page {page_num} of "{pdf_file}" and saved to "{excel_file}" successfully.')
                        
                        # Step 4: Extract and store additional information
                        # Example: Extract Heat Exchanger Name and Effective Area Value from tables
                        for table in tables:
                            # Assuming table.df is the DataFrame extracted from Camelot
                            # Example logic to extract specific values (modify as per your actual data)
                            for index, row in table.df.iterrows():
                                if 'Heat Exchanger Name' in row.values:
                                    pdf_data[pdf_file]['Heat Exchanger Name'] = row.values[1]  # Assuming name is in the second column
                                if 'Effective Area' in row.values:
                                    pdf_data[pdf_file]['Effective Area Value'] = row.values[1]  # Assuming value is in the second column
            else:
                print(f'The word "{search_word}" was not found in "{pdf_file}".')

    return pdf_data

if __name__ == "__main__":
    current_dir = os.path.dirname(os.path.abspath(__file__))
    folder_path = current_dir
    search_words = ["Effective Area"]
    
    process_all_pdfs_in_folder(folder_path, search_words)


# _______________________________


import os
import pandas as pd

# Assuming the Excel files are in the same directory as this script
current_dir = os.path.dirname(os.path.abspath(__file__))
excel_files = [file for file in os.listdir(current_dir) if file.endswith('.xlsx') or file.endswith('.xls')]

# Check if there are any Excel files in the current directory
if len(excel_files) > 0:
    for excel_file in excel_files:
        excel_file_path = os.path.join(current_dir, excel_file)
        
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(excel_file_path)
        
        # Print the name of the current Excel file being processed
        print(f"Processing Excel file: {excel_file}")
        
        # Search for the cell containing 'Service'
        service_value = None
        
        for i, row in df.iterrows():
            for j, cell in enumerate(row):
                if isinstance(cell, str) and 'Service' in cell:
                    service_value = df.iloc[i, j+2] or df.iloc[i, j+1] or df.iloc[i, j+3]
                    break
            if service_value is not None:
                break
        
        if service_value is None:
            print("Service not found.")
        else:
            print(f"Heat Exchanger Name: {service_value}")

        # Search for the cell containing 'Overall Area'
        overall_area_value = None
        
        for i, row in df.iterrows():
            for j, cell in enumerate(row):
                if isinstance(cell, str) and 'Effective Area' in cell:
                    # Attempt to find the value in the next few cells to the right
                    for offset in range(1, 10):  # Adjust the range based on how many cells to check
                        if j + offset < len(df.columns):
                            value = df.iloc[i, j + offset]
                            if pd.notna(value):
                                overall_area_value = value
                                break
                    break
            if overall_area_value is not None:
                break
        
        if overall_area_value is None:
            print("Effective Area not found.")
        else:
            print(f"Effective Area Value: {overall_area_value}")

        # Search for the cell containing 'Heat Duty'
        heat_duty_row = None
        heat_duty_col = None
        
        for i, row in df.iterrows():
            for j, cell in enumerate(row):
                if isinstance(cell, str) and 'Heat Duty' in cell:
                    heat_duty_row = i
                    heat_duty_col = j
                    break
            if heat_duty_row is not None:
                break
        
        if heat_duty_row is None:
            print("Heat Duty not found.")
        else:
            # Define the range around the cell containing 'Heat Duty'
            start_row = max(heat_duty_row - 2, 0)
            end_row = min(heat_duty_row + 1, len(df))
            start_col = max(heat_duty_col - 1, 0)
            end_col = min(j + 30, len(df.columns))
            
            # Display the surrounding cells, filtering out rows and columns that are entirely NaN
            surrounding_cells = df.iloc[start_row:end_row, start_col:end_col].dropna(how='all', axis=0).dropna(how='all', axis=1)
            
            # Set pandas display options to show all columns without truncation
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', None)
            
            print(surrounding_cells)
        
        print("______________________________________________________________________________________________________________________________")

else:
    print("No Excel files found in the folder.")
