<h1>EnigmaPDF: A PDF Management Tool</h1>
EnigmaPDF is a comprehensive application for handling various PDF operations. This GUI application, built with PySide6 and PyMuPDF (fitz), provides functionalities for converting, extracting, splitting, merging, encrypting, watermarking, and more. It integrates with Microsoft Word for document conversion and supports batch processing for multiple files.

Features
<h2>File Operations</h2>

Select File: Allows users to choose a single file for processing.
Select Folder: Enables batch operations by selecting a folder containing files.
Select Output Folder: Defines where the processed files will be saved.

<h3>File and Folder Management</h3>
Function: Manages files and folders for batch processing.

File Selection: Choose files individually using select_file or in bulk from a folder using select_folder.
Folder Selection: Specify output directories for saving processed files.
File Display: The show_files widget displays a list of selected files with their paths.
Steps:

Add files or folders to the processing list.
1.Select the output folder for saving results.
2.Manage the list of files and their display in the application.

Batch Processing:
Function: Handles batch processing for converting, extracting, and merging PDFs.

Batch File Addition: Use select_folder to add all files of a certain type from a directory.
Processing: Perform operations like conversion or extraction on all selected files.

Steps:
1.Use select_folder to add multiple files.
2.Perform the desired batch operation (e.g., conversion, extraction).
3.Each feature is designed to provide a user-friendly interface for performing common PDF operations.

Select File: Opens a dialog to choose a file, updates the file path, and manages page ranges for removal.
Select Folder: Opens a dialog to choose a folder and populates the file list based on file extensions.
Add Files: Displays selected files in a list.
Select Output Folder: Sets the output folder for processed files.
Select Document: Chooses a PDF document for insertion operations.

<h2>PDF Operations</h2>

Convert: Handles conversion between formats using Microsoft Word (for DOC/DOCX) and PyMuPDF (for image to PDF and PDF to image).
Extract: Extracts text or images from selected PDF files.
Split PDF: Divides a PDF into two files based on user-defined split points.
Merge PDF: Combines multiple PDFs into one.
Encrypt PDF: Adds password protection to PDFs.
Watermark PDF: Adds watermark text to each page of a PDF.
Remove Pages: Deletes specified pages from a PDF.
Insert Pages: Inserts pages from one PDF into another.

Here’s a detailed breakdown of each feature:

<h3>1. PDF Splitting</h3>
Function: Splits a PDF file into two separate documents at a specified page.

Page Selection: Use the page_dropdown to select the page at which the PDF will be split.
Output Files: The resulting split PDFs are saved in the selected output directory with names specified in the output1 and output2 text fields.
Usage: Ideal for dividing large documents into smaller, manageable files.
Steps:
1.Select a PDF file using the select_file button.
2.Enter the page number for splitting.
3.Click split_button to execute the split.

<h3>2. PDF Insertion</h3>
Function: Inserts one PDF into another at a specified location.

Document Selection: Use the select_document button to choose the PDF to be inserted.
Insert Options: Select between inserting the entire document, a single page, or a range of pages from the insert document.
Positioning: Specify the page in the original document where the insertion will occur.
Steps:
1.Select the main PDF and the PDF to be inserted using select_file and select_document.
2.Choose the insertion options and page ranges.
3.Click insert_button to perform the insertion.

<h3>3. PDF Conversion </h3>
Function: Converts files between different formats.

Supported Conversions:
Doc/Docx to PDF: Converts Microsoft Word documents to PDF.
Image to PDF: Converts images (JPG, PNG, JPEG) to PDF.
PDF to Image: Converts each page of a PDF into an image file (JPG, JPEG, PNG).
Steps:
1.Add files to be converted using select_file or select_folder.
2.Choose the conversion type from convert_dropdown.
3.Click convert_button to start the conversion.

<h3>4. Text and Image Extraction</h3>
Function: Extracts text or images from PDF files.

Text Extraction: Extracts all text from the selected PDFs and saves it to a text file.
Image Extraction: Extracts all images from the PDFs and saves them as separate image files.
Steps:
1.Add PDFs to extract text or images from using select_file or select_folder.
2.Select extraction type from extract_dropdown (Text or Images).
3.Click extract_button to perform the extraction.

<h3>5. PDF Merging</h3>
Function: Merges multiple PDF files into a single document.

File Selection: Use select_file or select_folder to choose PDFs to be merged.
Output: The merged PDF is saved in the selected output directory.
Steps:
1.Add PDFs to be merged.
2.Click merge_button to start the merging process.

<h3>6. PDF Encryption</h3>
Function: Encrypts PDF files with a password.

Password: Set a password for the PDF files to restrict access.
Output: Encrypted PDFs are saved in the selected output directory with a prefix "encrypted_".
Steps:
1.Set a password in the password_text field.
2.Add PDFs to be encrypted.
3.Click password_button to encrypt the PDFs.

<h3>7. Watermarking</h3>
Function: Adds a watermark text to each page of a PDF.

Text and Positioning: Enter the watermark text and choose its position (Top, Center, Bottom).
Output: Watermarked PDF is saved in the selected output directory with the name "watermarked_output.pdf".
Steps:

1.Enter the watermark text and choose its position.
2.Select the PDF to watermark.
3.Click watermark_button to apply the watermark.

Error Handling and Validation

Checks for required inputs and displays error messages if inputs are missing or invalid.
Provides feedback on successful operations or failures.

<h2>How to Use PdfWizard</h2>
<b>Setting Up</b>
Install Dependencies: Ensure you have Python and the required libraries installed. You may need PySide6, fitz, win32com, and win32api. Install them using pip if they are not already installed.

<b>Select File</b>

Click the Select File button to open a file dialog.
Choose a file to work with. The selected file’s path will be displayed in the UI.

<b>Add Folder for Batch Conversion</b>

1.Click the Select Folder button.
2.Choose a folder containing files you want to process in batch mode.
3.The application will add files to the list based on the file extensions and selected conversion option.
4. Or you can add Files one by one
5. Use the Add Files button to add files to the list.
6. Files will be displayed in the file list widget.

<b>Set Output Path</b>
1. Click the Select Output Folder button to specify where the processed files will be saved.

<b>Execute Operations</b>

1.Choose the operation you want to perform (e.g., Convert, Split, Merge) by selecting the corresponding button.
C2.onfigure any necessary options (e.g., page ranges, conversion formats) in the UI.
3.Click the operation button to execute the task.

<h3>Error Handling:</h3> The application will display error messages if operations fail or required inputs are missing.

<h3>Contributing</h3>
Feel free to fork the repository, make changes, and submit pull requests. For issues or feature requests, open a new issue on the GitHub repository.

<h3>License</h3>
This project is licensed under the MIT License. See the LICENSE file for details.
