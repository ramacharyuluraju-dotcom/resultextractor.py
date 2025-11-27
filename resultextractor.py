import pdfplumber
import pandas as pd
import re
import os

def extract_vtu_results(pdf_path):
    extracted_data = []
    
    # Check if file is actually a PDF
    if not pdf_path.lower().endswith('.pdf'):
        return []

    print(f"Scanning file: {os.path.basename(pdf_path)}...")

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                
                # 1. Extract Student Info
                # Regex designed to catch USN/Name even if they are on new lines
                usn_pattern = r"University Seat Number\s*[:,-]?\s*([0-9A-Z]+)"
                name_pattern = r"Student Name\s*[:,-]?\s*([A-Za-z\s]+)"
                
                usn_match = re.search(usn_pattern, text)
                name_match = re.search(name_pattern, text)
                
                usn = usn_match.group(1).strip() if usn_match else "Unknown"
                name = name_match.group(1).strip() if name_match else "Unknown"
                
                # 2. Extract Tables
                tables = page.extract_tables()
                
                for table in tables:
                    for row in table:
                        # Clean the row data (remove newlines inside cells)
                        clean_row = [str(item).replace('\n', ' ').strip() if item else '' for item in row]
                        
                        # Check for valid result row (starts with a Subject Code like BMATE201)
                        # We look for a code that is at least 5 alphanumeric characters
                        if len(clean_row) >= 6 and re.match(r'^[A-Z0-9]{5,}', clean_row[0]):
                            
                            sub_code = clean_row[0]
                            sub_name = clean_row[1]
                            internal = clean_row[2]
                            external = clean_row[3]
                            total = clean_row[4]
                            result = clean_row[5]

                            # --- FIX: Handle Merged Internal/External Marks ---
                            # Sometimes PDF tables merge cols 2 and 3 into Col 2 (e.g., "25 40")
                            # If External is empty and Internal has two numbers, split them.
                            if not external and ' ' in internal:
                                parts = internal.split()
                                if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                                    internal = parts[0]
                                    external = parts[1]

                            extracted_data.append({
                                "USN": usn,
                                "Student Name": name,
                                "Subject Code": sub_code,
                                "Subject Name": sub_name,
                                "Internal": internal,
                                "External": external,
                                "Total": total,
                                "Result": result,
                                "Source File": os.path.basename(pdf_path)
                            })
    except Exception as e:
        print(f"Error reading {os.path.basename(pdf_path)}: {e}")
        
    return extracted_data

def process_bulk_results(folder_path, output_excel):
    all_results = []
    files_found = False
    
    # Loop through all files in the CURRENT folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            files_found = True
            pdf_full_path = os.path.join(folder_path, filename)
            data = extract_vtu_results(pdf_full_path)
            all_results.extend(data)

    if not files_found:
        print("No PDF files found in this folder.")
        return

    # Export to Excel
    if all_results:
        df = pd.DataFrame(all_results)
        output_path = os.path.join(folder_path, output_excel)
        df.to_excel(output_path, index=False)
        print(f"\nSuccess! Extracted {len(all_results)} rows.")
        print(f"Data saved to: {output_excel}")
    else:
        print("PDFs found, but no matching result data could be extracted.")

if __name__ == "__main__":
    # Look in the SAME directory where the script is running
    current_dir = os.getcwd()
    output_file = "Bulk_VTU_Results.xlsx"
    
    print(f"Looking for PDFs in: {current_dir}")
    process_bulk_results(current_dir, output_file)