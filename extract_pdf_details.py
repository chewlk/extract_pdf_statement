import os
import re
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def extract_transactions():
    # Create a root window but hide it
    root = tk.Tk()
    root.withdraw()
    
    # Show the folder selection dialog
    pdf_folder = filedialog.askdirectory(title="Select folder containing bank statement PDFs")
    
    # If user cancels the dialog, exit the function
    if not pdf_folder:
        print("❌ Folder selection cancelled.")
        return
    
    # Define output file path
    output_file = os.path.join(pdf_folder, "extracted_data.xlsx")
    
    # Updated regex pattern to match BUS/MRT transactions
    transaction_pattern = re.compile(r"(\d{2} [A-Za-z]{3})\s+(\d{2} [A-Za-z]{3})\s+(BUS/MRT \d+)\s+SINGAPORE\s+([\d,.]+)", re.IGNORECASE)
    
    # List to store extracted rows
    all_rows = []
    
    # Count processed files and matches for reporting
    processed_files = 0
    
    # Loop through all PDF files
    for filename in os.listdir(pdf_folder):
        if filename.endswith(".pdf"):
            processed_files += 1
            filepath = os.path.join(pdf_folder, filename)
            print(f"Processing: {filename}")
            
            try:
                with pdfplumber.open(filepath) as pdf:
                    # Loop through each page
                    for page_num, page in enumerate(pdf.pages, start=1):
                        page_text = page.extract_text()
                        if page_text:
                            for line in page_text.split("\n"):
                                match = transaction_pattern.search(line)
                                if match:
                                    post_date, trans_date, description, amount = match.groups()
                                    all_rows.append([filename, post_date, trans_date, description, amount])
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")
    
    # Save extracted data to Excel
    if all_rows:
        df = pd.DataFrame(all_rows, columns=["PDF Filename", "Post Date", "Transaction Date", "Description", "SGD"])
        df.to_excel(output_file, index=False)
        print(f"✅ Data saved to: {output_file}")
        print(f"✅ Processed {processed_files} PDF files and extracted {len(all_rows)} transactions.")
    else:
        print("❌ No relevant rows found.")
        print(f"✅ Processed {processed_files} PDF files.")

if __name__ == "__main__":
    extract_transactions()