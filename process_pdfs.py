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
    
    # REGEX GUESS: Assumes "Bid No" (case-insensitive) followed by a space or colon
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
                    
                    # Search for bid number if not already found
                    if bid_num == "NOT FOUND":
                        bid_match = bid_pattern.search(text)
                        if bid_match:
                            bid_num = bid_match.group(1).strip()

                # Optimization: stop searching if we found both
                if cert_num != "NOT FOUND" and bid_num != "NOT FOUND":
                    break
                        
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return ("ERROR READING PDF", "ERROR READING PDF")
        
    return (cert_num, bid_num)

def main():
    """
    Main function to find PDFs, extract data, and write to a CSV.
    """
    folder_path = os.getcwd()  # Gets the current directory where the script is run
    output_filename = 'results.csv'
    all_data_rows = []
    
    # Headers for the CSV file
    headers = ["No.", "BID Number in Filename", "Cert", "Name", "Certificate Number", "Bid number"]
    
    file_counter = 1
    
    # Loop through all files in the current directory
    for file_name_raw in os.listdir(folder_path):
        # file_name = file_name_raw.strip() # Trim leading/trailing whitespace
        
        # Check the trimmed version of the file, but use the raw name for the path
        if file_name_raw.strip().lower().endswith('.pdf'):
            
            # 1. Remove .pdf extension from the *trimmed* name
            base_name = file_name_raw.strip()[:-4]
            
            # 2. Split file name by space
            parts = base_name.split(" ")
            
            # 3. Get the full path for the PDF reader
            #    Use the original (un-stripped) name as the file system needs it
            full_pdf_path = os.path.join(folder_path, file_name_raw)
            
            # 4. Extract both numbers from the PDF content
            (certificate_num, bid_num) = extract_pdf_data(full_pdf_path)
            
            # 5. Prepare the data row
            # If there are more than 2 parts, join all parts from index 2 onwards
            # with a space to form the "Name".
            name_part = " ".join(parts[2:]) if len(parts) > 2 else ''
            
            row = [
                file_counter,
                parts[0] if len(parts) > 0 else '',  # Part 1
                parts[1] if len(parts) > 1 else '',  # Part 2
                name_part,                           # Part 3 (and all subsequent parts)
                certificate_num,
                bid_num
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


