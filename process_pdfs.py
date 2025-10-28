import os
import re
import csv
import pdfplumber

def extract_pdf_data(pdf_path):
    """
    Opens a PDF and searches all pages for the certificate and bid numbers.
    Returns a tuple: (certificate_number, bid_number).
    """
    # This regex looks for "Certificate number:" (case-insensitive)
    # and captures the alphanumeric string (and hyphens) that follows.
    cert_pattern = re.compile(r"Certificate Number :\s*([\w-]+)", re.IGNORECASE)
    
    # <<< CHANGED
    # REGEX GUESS: Assumes "Bid Number" followed by a space or colon
    # You may need to change "Bid Number" to match your PDF exactly.
    bid_pattern = re.compile(r"Bid No[\s:]*([\w-]+)", re.IGNORECASE)
    
    cert_num = "NOT FOUND"
    bid_num = "NOT FOUND"
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Extract text from the current page
                text = page.extract_text()
                
                if text:
                    # Search for cert number if not already found
                    if cert_num == "NOT FOUND":
                        cert_match = cert_pattern.search(text)
                        if cert_match:
                            cert_num = cert_match.group(1).strip()
                    
                    # <<< CHANGED: Search for bid number if not already found
                    if bid_num == "NOT FOUND":
                        bid_match = bid_pattern.search(text)
                        if bid_match:
                            bid_num = bid_match.group(1).strip()

                # Optimization: stop searching if we found both
                if cert_num != "NOT FOUND" and bid_num != "NOT FOUND":
                    break
                        
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return ("ERROR READING PDF", "ERROR READING PDF") # <<< CHANGED
        
    return (cert_num, bid_num) # <<< CHANGED

def main():
    """
    Main function to find PDFs, extract data, and write to a CSV.
    """
    folder_path = os.getcwd()  # Gets the current directory where the script is run
    output_filename = 'results.csv'
    all_data_rows = []
    
    # Headers for the CSV file
    headers = ["No.", "BID Number in Filename", "Cert", "Name", "Certificate Number", "Bid number"] # <<< CHANGED
    
    file_counter = 1
    
    # Loop through all files in the current directory
    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith('.pdf'):
            
            # 1. Remove .pdf extension
            base_name = file_name[:-4]
            
            # 2. Split file name by space
            parts = base_name.split(" ")
            
            # 3. Get the full path for the PDF reader
            full_pdf_path = os.path.join(folder_path, file_name)
            
            # 4. Extract both numbers from the PDF content
            # <<< CHANGED
            (certificate_num, bid_num) = extract_pdf_data(full_pdf_path)
            
            # 5. Prepare the data row
            # <<< CHANGED
            row = [
                file_counter,
                parts[0] if len(parts) > 0 else '',  # Part 1
                parts[1] if len(parts) > 1 else '',  # Part 2
                parts[2] if len(parts) > 2 else '',  # Part 3
                certificate_num,
                bid_num  # Added the new bid number
            ]
            
            all_data_rows.append(row)
            file_counter += 1
            
    # 6. Write all collected data to the CSV file
    try:
        with open(output_filename, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(headers)
            writer.writerows(all_data_rows)
            
        print(f"Success! Processed {file_counter - 1} PDF files.")
        print(f"Data saved to: {output_filename}")
        
    except PermissionError:
        print(f"\n--- ERROR ---")
        print(f"Could not write to {output_filename}.")
        print(f"Please close the file in Excel and run the script again.")
    except Exception as e:
        print(f"\n--- ERROR ---")
        print(f"An unexpected error occurred while writing the CSV: {e}")

if __name__ == "__main__":
    main()