# **PDF Certificate & Bid Number Extractor**

This is a Python script that scans all PDF files in a folder, extracts a **Certificate Number** and a **Bid Number** from each, and saves all the collected data into a single results.csv file.

The script is designed to run in its current directory, making it easy to use: just place it (or the compiled .exe) in the folder with your PDFs and run it.

## **Features**

* **Batch Processing:** Processes all .pdf files in the same folder.  
* **Data Extraction:** Uses regular expressions to find:  
  * Certificate Number (e.g., Certificate Number : 123-ABC)  
  * Bid Number (e.g., Bid No: 456-DEF)  
* **CSV Output:** Creates a results.csv file with the following columns:  
  * No. (A running counter)  
  * BID Number in Filename (Part 1 of the filename)  
  * Cert (Part 2 of the filename)  
  * Name (Part 3 of the filename)  
  * Certificate Number (Data found *inside* the PDF)  
  * Bid number (Data found *inside* the PDF)  
* **Error Handling:**  
  * Gracefully handles PDFs it can't read and logs "ERROR READING PDF" in the CSV.  
  * Warns you if the results.csv file is open in Excel (PermissionError) so you can close it and try again.

## **How to Use (For End-Users)**

This is the easiest way to use the tool. You don't need Python installed for this.

1. Go to the [**Releases**](https://www.google.com/search?q=https://github.com/YourUsername/YourRepo/releases) page of this repository.  
2. Download the latest process\_pdfs.exe file.  
3. **Copy** the process\_pdfs.exe file into the folder that contains all the PDF files you want to process.  
4. **Double-click process\_pdfs.exe** from inside that folder.  
5. A console window will appear, and after a moment, it will show a "Success\!" message.  
6. Look for a new file named results.csv in that same folder. This file contains all your extracted data.

## **For Developers**

### **Requirements**

* Python 3.6+  
* pdfplumber  
* pyinstaller (for building)

### **Installation**

1. Clone this repository:  
   git clone \[https://github.com/YourUsername/YourRepo.git\](https://github.com/YourUsername/YourRepo.git)  
   cd YourRepo

2. (Recommended) Create a virtual environment:  
   python \-m venv venv  
   .\\venv\\Scripts\\activate  \# On Windows  
   \# source venv/bin/activate  \# On macOS/Linux

3. Install the required libraries:  
   pip install pdfplumber

   *(Or, if a requirements.txt is provided: pip install \-r requirements.txt)*

### **Running the Script**

1. Place the process\_pdfs.py script in the folder with your PDFs.  
2. Run it directly from the terminal:  
   python process\_pdfs.py

3. The results.csv file will be created in that folder.

### **How to Build the .EXE Yourself**

If you've made changes to the script and want to re-compile the executable:

1. Make sure you have PyInstaller installed:  
   pip install pyinstaller

2. From your terminal, in the repository's main folder, run:  
   pyinstaller \--onefile process\_pdfs.py

3. PyInstaller will run and create two new folders (build and dist).  
4. Your new executable file will be located in the dist folder, named **process\_pdfs.exe**.