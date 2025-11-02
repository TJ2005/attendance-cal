import PyPDF2
import pdfplumber
import pandas as pd

# Path to your PDF file
pdf_path = "ZSVKM_STUDENT_ATTENDANCE.pdf"

print("=" * 80)
print("ANALYZING ATTENDANCE PDF")
print("=" * 80)

# Method 1: Using PyPDF2 to extract raw text
print("\n--- METHOD 1: PyPDF2 Raw Text Extraction ---\n")
try:
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        print(f"Total pages: {len(pdf_reader.pages)}\n")
        
        for page_num in range(len(pdf_reader.pages)):
            print(f"\n{'='*60}")
            print(f"PAGE {page_num + 1}")
            print(f"{'='*60}\n")
            page = pdf_reader.pages[page_num]
            text = page.extract_text()
            print(text)
            print("\n")
except Exception as e:
    print(f"Error with PyPDF2: {e}")

# Method 2: Using pdfplumber for better table extraction
print("\n" + "=" * 80)
print("--- METHOD 2: pdfplumber Table Extraction ---")
print("=" * 80 + "\n")
try:
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            print(f"\n{'='*60}")
            print(f"PAGE {page_num + 1} - Detailed Analysis")
            print(f"{'='*60}\n")
            
            # Extract text
            text = page.extract_text()
            print("Raw Text:")
            print(text)
            print("\n")
            
            # Try to extract tables
            tables = page.extract_tables()
            if tables:
                print(f"Found {len(tables)} table(s) on this page:\n")
                for i, table in enumerate(tables):
                    print(f"\n--- Table {i + 1} ---")
                    print(f"Rows: {len(table)}, Columns: {len(table[0]) if table else 0}\n")
                    
                    # Print table structure
                    for row_idx, row in enumerate(table):
                        print(f"Row {row_idx}: {row}")
                    
                    # Try to create a DataFrame for better visualization
                    if len(table) > 1:
                        try:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            print("\n--- DataFrame Preview ---")
                            print(df.head(20))
                            print(f"\nShape: {df.shape}")
                            print(f"Columns: {df.columns.tolist()}")
                        except Exception as e:
                            print(f"Could not create DataFrame: {e}")
            else:
                print("No tables detected on this page.")
            
            print("\n" + "-" * 60)
            
except Exception as e:
    print(f"Error with pdfplumber: {e}")

print("\n" + "=" * 80)
print("ANALYSIS COMPLETE")
print("=" * 80)
