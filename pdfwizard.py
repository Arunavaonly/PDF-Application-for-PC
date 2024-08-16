from ui import Ui_MainWindow
from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PySide6.QtCore import QTimer
import os
import win32com.client
import win32api
import fitz


class PdfWizard(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("EnigmaPDF")

        self.mainremove_button.clicked.connect(self.switch_removepage)
        self.mainconvert_button.clicked.connect(self.switch_convertpage)
        self.mainextract_button.clicked.connect(self.switch_extractpage)
        self.mainsplit_button.clicked.connect(self.switch_splitpage)
        self.maininsert_button.clicked.connect(self.switch_insertpage)
        self.mainmerge_button.clicked.connect(self.switch_mergepage)
        self.mainpassword_button.clicked.connect(self.switch_protectpage)
        self.mainwatermark_button.clicked.connect(self.switch_watermarkpage)

        self.select_button.clicked.connect(self.select_file)
        self.addfiles_button.clicked.connect(self.add_file)
        self.select_button.clicked.connect(self.show_page_range)
        self.select_button.clicked.connect(self.show_page_range_insert)
        self.selectfolder_button.clicked.connect(self.select_folder)
        self.output_button.clicked.connect(self.select_output_folder)
        self.select_document_insert.clicked.connect(self.select_document)

        self.split_button.clicked.connect(lambda: self.blink_and_execute(self.split_button, self.split_pdf))

        self.insert_button.clicked.connect(lambda: self.blink_and_execute(self.insert_button, self.insert_pdf))

        self.convert_button.clicked.connect(lambda: self.blink_and_execute(self.convert_button, self.convert))

        self.extract_button.clicked.connect(lambda: self.blink_and_execute(self.extract_button, self.extract))

        self.merge_button.clicked.connect(lambda: self.blink_and_execute(self.merge_button, self.merge_pdf))
        
        self.password_button.clicked.connect(lambda: self.blink_and_execute(self.password_button, self.encrypt_pdf))

        self.remove_button.clicked.connect(lambda: self.blink_and_execute(self.remove_button, self.remove))

        self.watermark_button.clicked.connect(lambda: self.blink_and_execute(self.watermark_button, self.watermark))

        self.selected_file_path = ""
        self.selected_output_path = ""
        self.document_for_insert = ""
        self.file_list = []

        self.select_document_widget.setVisible(False)

        # Ensure convertPage is shown and mainConvert is checked
        self.mainconvert_button.setChecked(True)
        self.stackedWidget.setCurrentWidget(self.convert_page)

    def blink_button(self, button):
        # Hide the button
        button.setVisible(False)
        # Create a QTimer to show the button after a few milliseconds
        QTimer.singleShot(350, lambda: button.setVisible(True))

    def blink_and_execute(self, button, func):
        self.blink_button(button)  # Make the button blink
        func()  # Execute the button's associated function

        
    def set_all_enabled(self, enabled):
        self.addfiles_button.setEnabled(enabled)
        self.selectfolder_button.setEnabled(enabled)


    def get_file_filter(self):
        # Check which button is checked and return the appropriate file filter
        if self.mainconvert_button.isChecked():
            if self.convert_dropdown.currentText() == "Doc/Docx to PDF":
                return "Word Documents (*.doc *.docx)"
            elif self.convert_dropdown.currentText() == "Image to PDF":
                return "Image Files (*.png *.jpg *.jpeg )"
            # Add other conditions based on your UI logic'''
        else:
            return "PDF Files (*.pdf)"

    def select_file(self):
        self.file_path.clear()  
        self.selected_file_path = ""  
        file_filter = self.get_file_filter() 
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", file_filter, options=options)  # Modified to use the correct file filter (Line 8)
        if file_path:  # Check if a file was selected (Line 9)
            self.selected_file_path = file_path
            self.file_path.setText(file_path)
            self.file_list.append(file_path)
            if self.mainremove_button.isChecked() and self.selected_file_path:  # Check if remove button is checked (Line 13)
                document = fitz.open(file_path)
                pages = document.page_count
                self.page_range_remove.setText(f"Page Range: 1-{pages}")
                self.start_page_dropdown_remove.clear()
                self.end_page_dropdown_remove.clear()
                for page_number in range(1, pages+1):
                    self.start_page_dropdown_remove.addItem(str(page_number))
                    self.end_page_dropdown_remove.addItem(str(page_number))

    #document selection for insert_pdf function
    def select_document(self):
        options = QFileDialog.Options()
        document_selected,_ = QFileDialog.getOpenFileName(self, "Select File", "", "PDF Files (*.pdf)", options=options)
        if document_selected:
            self.document_for_insert = document_selected
            self.insert_document_path.setText(document_selected)
            document = fitz.open(document_selected)
            insert_range = document.page_count
            self.start_page_dropdown.clear()
            self.end_page_dropdown.clear()
            for page in range(1, insert_range + 1):
                self.start_page_dropdown.addItem(str(page))
                self.end_page_dropdown.addItem(str(page))
   

    def select_folder(self):
        self.folder_path.clear()  # Clear the folder path (Line 3)
        folder_path = QFileDialog.getExistingDirectory(None, "Select Folder", "", QFileDialog.ShowDirsOnly)

        if folder_path:
            self.folder_path.setText(folder_path)
            self.file_list.clear()  # Clear file_list before adding new files (Line 8)
            if self.mainconvert_button.isChecked() and self.convert_dropdown.currentText() == "Doc/Docx to PDF":  # Modified to add files based on conversion option (Lines 10-16)
                for file in os.listdir(folder_path):
                    if file.lower().endswith(".doc") or file.lower().endswith(".docx"):
                        full_path = os.path.join(folder_path, file).replace("\\", '/')
                        self.file_list.append(full_path)
            elif self.mainconvert_button.isChecked() and self.convert_dropdown.currentText() == "Image to PDF":  # Modified to add files based on conversion option (Lines 17-23)
                for file in os.listdir(folder_path):
                    if file.lower().endswith(".jpg") or file.lower().endswith(".jpeg") or file.lower().endswith(".png"):
                        full_path = os.path.join(folder_path, file).replace("\\", '/')
                        self.file_list.append(full_path)
            else:
                for file in os.listdir(folder_path):
                    if file.lower().endswith(".pdf"):
                        full_path = os.path.join(folder_path, file).replace("\\", '/')
                        self.file_list.append(full_path)

    def add_file(self):
        self.show_files.clear()
        for index, file in enumerate(self.file_list):
            self.show_files.append(f"{index+1}.{file}\n")

    def select_output_folder(self):
        folder_path = QFileDialog.getExistingDirectory(None, "Select Folder", "", QFileDialog.ShowDirsOnly)
        if folder_path:
            self.selected_output_path = folder_path
            # Update the line edit with the selected folder path
            self.output_path.setText(folder_path)

    def insert_pdf(self):
        if not self.selected_file_path or not self.selected_output_path or not self.document_for_insert:
            QMessageBox.warning(self, "Input Error", "Please provide all required fields")
            return
        try:
            pdf = fitz.open(self.selected_file_path)
            document = fitz.open(self.document_for_insert)
            insert_position = int(self.insert_at_dropdown.currentText())
            start = int(self.start_page_dropdown.currentText())
            end = int(self.end_page_dropdown.currentText())

            new_pdf = fitz.open()

            if self.Insert_options_dropdown.currentText() == 'Whole Document':
                new_pdf.insert_pdf(pdf, from_page=0, to_page=insert_position -1)
                new_pdf.insert_pdf(document, from_page=0, to_page=len(document) - 1)
                new_pdf.insert_pdf(pdf, from_page=insert_position, to_page=len(pdf) - 1)
            elif self.Insert_options_dropdown.currentText() == "Signle Page":
                new_pdf.insert_pdf(pdf, from_page=0, to_page=insert_position -1)
                new_pdf.insert_pdf(document, from_page=start -2, to_page=start -1)
                new_pdf.insert_pdf(pdf, from_page=insert_position, to_page=len(pdf) - 1)
            else:
                new_pdf.insert_pdf(pdf, from_page=0, to_page=insert_position -1)
                new_pdf.insert_pdf(document, from_page=start -2, to_page=end -1)
                new_pdf.insert_pdf(pdf, from_page=insert_position, to_page=len(pdf) - 1)
            
            output = os.path.join(self.selected_output_path, "inserted.pdf")
            new_pdf.save(output)
            new_pdf.close()
            pdf.close()
            document.close()
            self.statusBar().showMessage(f"Success: Insert action completed successfully.", 1000)
            QApplication.processEvents()
            self.statusBar().showMessage("inserted.pdf saved", 1000)

            self.file_path.clear()
            self.insert_document_path.clear()
            self.selected_file_path = ""
            self.document_for_insert = ""
            self.page_range_insert.setText("Page Range: ")
            self.insert_at_dropdown.clear()
            self.start_page_dropdown.clear()
            self.end_page_dropdown.clear()

        except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occured while inserting: {e}")

    def show_page_range_insert(self):
        if self.maininsert_button.isChecked() and self.selected_file_path:
            pdf_file = self.selected_file_path
            document = fitz.open(pdf_file)
            pages = document.page_count
            self.page_range_insert.setText(f"Page Range: 1-{pages}")
            self.insert_at_dropdown.clear()
            for page_number in range(1, pages+1):
                self.insert_at_dropdown.addItem(str(page_number))


    def text_extract(self):
        for file in self.file_list:
            filename = os.path.basename(file)
            try:
            # Open the PDF file
                pdf = fitz.open(file)  # Replace with your PDF file path

                # Iterate over each page in the PDF
                for page_num in range(len(pdf)):
                    page = pdf[page_num]  # Get the page
                    text = page.get_text()  # Extract text from the page
                    text_file = os.path.join(self.selected_output_path,"extracted_text.txt")
                    with open(text_file, 'a', encoding = 'utf-8') as myfile:
                        myfile.write(text)
                # Close the PDF file
                pdf.close()
                self.statusBar().showMessage(f"Success: Text extracted from {filename} successfully.", 1000)
                QApplication.processEvents()

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to extract images: {e}")
                continue


    def image_extract(self):
        # Open the PDF file
        for file in self.file_list:
            filename = os.path.basename(file)
            try:
                pdf = fitz.open(file)  # Replace with your PDF file path
                # Iterate over each page in the PDF
                for page_num in range(len(pdf)):
                    page = pdf[page_num]  # Get the page
                    # Get the list of image dictionaries on the current PDF page
                    image_list = page.get_images(full=True)
                    if image_list:
                    # Iterate over the images on the page
                        for image_index, img in enumerate(image_list, start=1):
                            try:
                                xref = img[0]
                                base_image = pdf.extract_image(xref)
                                image_bytes = base_image["image"]
                                # Save the image to a file
                                image_file = os.path.join(self.selected_output_path, f'image_{page_num + 1}_{image_index}.png')
                                with open(image_file, 'wb') as img_file:
                                    img_file.write(image_bytes)
                            except Exception as e:
                                continue  # Skip this image and move to the next one
                    # Close the PDF file
                pdf.close()
                self.statusBar().showMessage(f"Success: Images extracted from {filename} successfully.", 1000)
                QApplication.processEvents()

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to extract images: {e}")
                continue

    def extract(self):
        # Get the selected option from the dropdown        
        if len(self.file_list) ==0:
            QMessageBox.warning(self, "Input Error", "Please select a PDF file.")
            return
        
        if self.selected_output_path:
            if self.extract_dropdown.currentText() == "Text":
                self.text_extract()

            elif self.extract_dropdown.currentText() == "Images":
                self.image_extract()
            else:
                QMessageBox.warning(self, "Input Error", "Please select a valid extraction option.")
                return
        else:
            QMessageBox.warning(self,  "No Output Path", "Please select an output path.")
            return

            
        self.show_files.clear()
        self.file_path.clear()
        self.folder_path.clear()

        self.file_list = []
        self.selected_file_path = ''
    
        
    def split_pdf(self):
        input_pdf = self.selected_file_path
        try:
            pdf = fitz.open(input_pdf)  # Open the PDF file
            doc1 = fitz.open()  # Create a new PDF for the first part
            doc2 = fitz.open()  # Create a new PDF for the second part
            #fetching the page number from the dropdown as selected bu the user
            split_page = int(self.page_dropdown.currentText())
            # Add pages to the first part (up to the split page)
            doc1.insert_pdf(pdf, from_page=0, to_page=split_page-1)

            # Add pages to the second part (from the split page to the end)
            doc2.insert_pdf(pdf, from_page=split_page)

            dir_path = os.path.split(input_pdf)[0]

            # Save the two parts
            if self.selected_output_path:
                output_pdf1 = os.path.join(self.selected_output_path, self.output1.text() +".pdf")
                doc1.save(output_pdf1)
                output_pdf2 = os.path.join(self.selected_output_path, self.output2.text() + ".pdf")
                doc2.save(output_pdf2)
            else:
                output_pdf1 = os.path.join(dir_path, self.output1.text() +".pdf")
                doc1.save(output_pdf1)
                output_pdf2 = os.path.join(dir_path, self.output2.text() + ".pdf")
                doc2.save(output_pdf2)

            # Close the documents
            pdf.close()
            doc1.close()
            doc2.close()
            self.statusBar().showMessage("Success: PDF splitted into two parts successfully.", 1000)

            self.selected_file_path = ""
            self.file_path.clear()
            self.page_dropdown.clear()
            self.page_range.setText("Page Range: ")

            QApplication.processEvents()

        except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to split the PDF: {e}")


    def show_page_range(self):
        if self.mainsplit_button.isChecked() and self.selected_file_path:
            pdf_file = self.selected_file_path
            document = fitz.open(pdf_file)
            pages = document.page_count
            self.page_range.setText(f"Page Range: 1-{pages}")
            self.page_dropdown.clear()
            for page_number in range(1, pages+1):
                self.page_dropdown.addItem(str(page_number))



    def encrypt_pdf(self):
        password = self.password_text.text()
        output_path = self.selected_output_path
        # Validate the inputs
        if len(self.file_list) == 0 or not output_path or not password:
            QMessageBox.warning(self, "Input Error", "Please provide all required inputs.")
            return
        try:
            for file in self.file_list:
                try:
                    # Open the input PDF file
                    document = fitz.open(file)

                    # Construct the output file path
                    base_name = os.path.basename(file)
                    output_file_path = os.path.join(output_path, f"encrypted_{base_name}")

                    # Encrypt the PDF with the specified password
                    document.save(output_file_path, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=password, user_pw=password)

                    # Close the document
                    document.close()
                    self.statusBar().showMessage(f'encrypted_{base_name} saved', 1000)
                    QApplication.processEvents()

                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to encrypt PDF: {e}")
                    continue
            self.statusBar().showMessage("Success: PDFs encrypted successfully.", 3000)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to encrypt PDF: {e}")

        # Clear the paths and file list
        self.selected_file_path = ""
        self.file_list = []

        self.show_files.clear()
        self.file_path.clear()
        self.folder_path.clear()
        

    def watermark(self):
        watermark = self.watermark_text.text()
        output_path = self.selected_output_path
        if self.selected_file_path == "" or not output_path or not watermark:
            QMessageBox.warning(self, "Input Error", "Please provide all required inputs.")
            return
        try:
            # Open the PDF file
            pdf = fitz.open(self.selected_file_path)  # Replace with your PDF file path

            # Define the font size and color
            font_size = 20
            font_color = (0.5, 0.5, 0.5)  # Grey color in RGB

            # Iterate over each page in the PDF
            for page_num in range(len(pdf)):
                page = pdf[page_num]  # Get the page
                # Calculate the position for the watermark
                text_rect = page.rect  # Get the page dimensions
                x_position = text_rect.width*0.42

                if self.watermark_position.currentText() == "Top":
                    y_position = text_rect.height*0.05
                elif self.watermark_position.currentText() == "Center":
                    y_position = text_rect.height*0.42
                else:
                    y_position = text_rect.height*0.95

                # Add the watermark text to the page
                page.insert_text(
                    (x_position, y_position),  # Position at the center of the page
                    watermark,  # The text to add
                    fontname='helv',  # Font
                    fontsize=font_size,  # Font size
                    color=font_color,  # Font color
                )
            # Save the watermarked PDF
            output_file = os.path.join(output_path, "watermarked_output.pdf")
            pdf.save(output_file)  # Replace with your desired output PDF file name

            # Close the PDF file
            pdf.close()
            self.statusBar().showMessage("Success: PDF watermarked successfully.", 1000)
            self.statusBar().showMessage("watermarked_output.pdf saved", 1000)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to watermark PDF: {e}")

        self.selected_file_path = ""
        self.file_path.clear()


    def merge_pdf(self):
        if len(self.file_list) ==0 or self.selected_output_path == "":
            QMessageBox.warning(self, "Input Error", "Please provide all required inputs.")
            return
        if len(self.file_list) ==1:
            QMessageBox.warning(self, "Single File Error", "Please provide at least two PDFs to start merging")
            return

        merged = fitz.open()
        for file in self.file_list:
            try:
                pdf_document = fitz.open(file)
                merged.insert_pdf(pdf_document)
                pdf_document.close()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to merged PDFs: {e}")
                continue

        if self.selected_output_path:
            merged_pdf = os.path.join(self.selected_output_path, "merged_output.pdf").replace("\\", '/')
            merged.save(merged_pdf)
            merged.close()
            self.statusBar().showMessage("Success: PDFs merged successfully.", 1000)
            self.statusBar().showMessage("merged_output.pdf saved", 1000)
        else:
            # Show warning message and prompt for output folder
            QMessageBox.warning(self, "No Output Path", "Please select an output path.")            
        
        self.show_files.clear()
        self.file_list = []
        self.selected_output_path = ""
        self.selected_file_path = ''

    

    def convert(self):
        if len(self.file_list) ==0:
            QMessageBox.warning(self, "Input Error", "Please select a file.")
            return
        
        if self.selected_output_path:
            if self.convert_dropdown.currentText() == "Doc/Docx to PDF":
                self.doc_to_pdf()
            elif self.convert_dropdown.currentText() == "Image to PDF":
                self.img2pdf()
            elif self.convert_dropdown.currentText() == "Pdf to Image(.jpg)":
                self.pdf2img("jpg")
            elif self.convert_dropdown.currentText() == "Pdf to Image(.jpeg)":
                self.pdf2img("jpeg")
            elif self.convert_dropdown.currentText() == "Pdf to Image(.png)":
                self.pdf2img("png")
            else:
                QMessageBox.warning(self, "Input Error", "Please select a valid conversion option.")
                return
        else:
            QMessageBox.warning(self, "No OutPut Path", "Please select an output path.")
            return
        
        self.show_files.clear()
        self.file_path.clear()
        self.folder_path.clear()

        self.file_list = []
        self.selected_file_path = ''

    def pdf2img(self, img_format):
        for file in self.file_list:
            filename = os.path.basename(file)
            try:
                pdf = fitz.open(file)
                # Iterate over each page in the PDF
                for page_num in range(len(pdf)):
                    page = pdf[page_num]  # Get the page
                    pix = page.get_pixmap()  # Render page to an image
                    pix.save(os.path.join(self.selected_output_path,f'output_page_{page_num + 1}.{img_format}'))  # Save the image
                self.statusBar().showMessage(f"Success: pages saved as images from {filename} successfully.", 1000)
                QApplication.processEvents()
                pdf.close()  # Close the PDF
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error: {e}")
                continue


    def doc_to_pdf(self):
        wdFormatPDF = 17
        # Start the Word application
        try:
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False  # Ensure Word is not visible during conversion
        except Exception as e:
            return
        
        try:
            for file in self.file_list:
                try:
                    filename = os.path.split(file)[1]
                    dir_path = os.path.split(file)[0]

                    # Convert the file path to the short path format
                    short_path = win32api.GetShortPathName(file)
                    doc = word.Documents.Open(short_path)
                    pdf_file = os.path.splitext(short_path)[0] + ".pdf"
                    pdf_file_path = os.path.join(self.selected_output_path, pdf_file)
                    doc.SaveAs(pdf_file_path, FileFormat=wdFormatPDF)
                    doc.Close()

                    new_filename = os.path.splitext(filename)[0] + ".pdf"
                    new_pdf_path = os.path.join(self.selected_output_path, new_filename)
                    os.rename(pdf_file_path, new_pdf_path)
                    self.statusBar().showMessage(f"Success: {filename} converted to PDF.", 1000)
                    QApplication.processEvents()
                except Exception as e:
                    QApplication.processEvents()
                    continue
        finally:
            word.Quit()


    def img2pdf(self):
        if len(self.file_list) == 0 or not self.selected_output_path:
            QMessageBox.warning(self, "Input Error", "Please select a file.")
            return
        try:
            doc = fitz.open()  # Create a new PDF
            for file in self.file_list:
                filename = os.path.basename(file)
                try:
                    img = fitz.open(file)  # Open the image
                    rect = img[0].rect  # Get the dimensions of the image
                    pdfbytes = img.convert_to_pdf()  # Convert the image to PDF bytes
                    img.close()  # Close the image file
                    imgpdf = fitz.open("pdf", pdfbytes)  # Open the image as a PDF
                    page = doc.new_page(width=rect.width, height=rect.height)  # Create a new page with the dimensions of the image
                    page.show_pdf_page(rect, imgpdf, 0)  # Insert the image PDF into the page
                    self.statusBar().showMessage(f"Success: {filename} converted into PDF successfully.", 1000)
                    QApplication.processEvents()
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Error: {e}")
                    continue
            output_pdf = os.path.join(self.selected_output_path, "output.pdf")     
            doc.save(output_pdf)  # Save the new PDF
            doc.close()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error: {e}")

    def remove(self):
        input_file = self.selected_file_path
        output_path = self.selected_output_path
        if output_path:
            try:
                output_file = os.path.join(self.selected_output_path, "pages_deleted.pdf")
                if input_file:
                    file_handle = fitz.open(input_file)
                    start = int(self.start_page_dropdown_remove.currentText()) #Start Index
                    end = int(self.end_page_dropdown_remove.currentText())    # End index
                    if not end - start< file_handle.page_count -1:
                        QMessageBox.warning(self,"Issue with Page Range", "Please change the page range")
                    else:
                        file_handle.delete_pages(start -1,end)
                        file_handle.save(output_file)
                        self.statusBar().showMessage(f"Success: pages removed and output file saved as pages_deleted.pdf.", 1000)
                else:
                    QMessageBox.warning(self, "No File Selected", "Please select a File.")

            except Exception as e:
                QMessageBox.critical(self, "Error", f'An error occured{e}')
        else:
            QMessageBox.warning(self, "No Output Path", "Please select an output path.")

        self.selected_file_path = ""

        self.start_page_dropdown_remove.clear()
        self.end_page_dropdown_remove.clear()
        self.page_range_remove.setText("Page Range")
        self.file_path.clear()

       

    def switch_removepage(self):
        self.folder_path.clear()
        self.show_files.clear()
        self.file_path.clear()
        self.output_path.clear()
        self.page_dropdown.clear()
        self.page_range.setText("Page Range: ")


        self.selected_file_path = ""
        self.selected_output_path = ""
        self.file_list = []

        self.stackedWidget.setCurrentWidget(self.remove_page)
        self.set_all_enabled(False)
        self.select_document_widget.setVisible(False)



    def switch_convertpage(self):
        self.folder_path.clear()
        self.show_files.clear()
        self.file_path.clear()
        self.output_path.clear()
        self.page_dropdown.clear()
        self.page_range.setText("Page Range: ")

        self.selected_file_path = ""
        self.selected_output_path = ""
        self.file_list = []


        self.stackedWidget.setCurrentWidget(self.convert_page)
        self.set_all_enabled(True)
        self.select_document_widget.setVisible(False)



    def switch_extractpage(self):
        self.folder_path.clear()
        self.show_files.clear()
        self.file_path.clear()
        self.output_path.clear()
        self.page_dropdown.clear()
        self.page_range.setText("Page Range: ")


        self.selected_file_path = ""
        self.selected_output_path = ""
        self.file_list = []

        self.stackedWidget.setCurrentWidget(self.extract_page)
        self.set_all_enabled(True)
        self.select_document_widget.setVisible(False)


    def switch_splitpage(self):
        self.folder_path.clear()
        self.show_files.clear()
        self.file_path.clear()
        self.output_path.clear()
        self.page_dropdown.clear()
        self.page_range.setText("Page Range: ")


        self.selected_file_path = ""
        self.selected_output_path = ""
        self.file_list = []

        self.stackedWidget.setCurrentWidget(self.split_page)
        self.selectfolder_button.setEnabled(False)
        self.addfiles_button.setEnabled(False)
        self.select_document_widget.setVisible(False)



    def switch_insertpage(self):
        self.folder_path.clear()
        self.show_files.clear()
        self.file_path.clear()
        self.output_path.clear()
        self.page_dropdown.clear()
        self.page_range.setText("Page Range: ")


        self.selected_file_path = ""
        self.selected_output_path = ""
        self.file_list = []

        self.stackedWidget.setCurrentWidget(self.insert_page)
        self.selectfolder_button.setEnabled(False)
        self.addfiles_button.setEnabled(False)
        self.select_document_widget.setVisible(True)



    def switch_mergepage(self):
        self.folder_path.clear()
        self.show_files.clear()
        self.file_path.clear()
        self.output_path.clear()
        self.page_dropdown.clear()
        self.page_range.setText("Page Range: ")


        self.selected_file_path = ""
        self.selected_output_path = ""
        self.file_list = []

        self.stackedWidget.setCurrentWidget(self.merge_page)
        self.set_all_enabled(True)
        self.select_document_widget.setVisible(False)


    def switch_protectpage(self):
        self.folder_path.clear()
        self.show_files.clear()
        self.file_path.clear()
        self.output_path.clear()
        self.page_dropdown.clear()
        self.page_range.setText("Page Range: ")


        self.selected_file_path = ""
        self.selected_output_path = ""
        self.file_list = []

        self.stackedWidget.setCurrentWidget(self.protect_page)
        self.set_all_enabled(True)
        self.select_document_widget.setVisible(False)


    def switch_watermarkpage(self):
        self.folder_path.clear()
        self.show_files.clear()
        self.file_path.clear()
        self.output_path.clear()
        self.page_dropdown.clear()
        self.page_range.setText("Page Range: ")


        self.selected_file_path = ""
        self.selected_output_path = ""
        self.file_list = []

        self.stackedWidget.setCurrentWidget(self.watermark_page)
        self.selectfolder_button.setEnabled(False)
        self.addfiles_button.setEnabled(False)
        self.select_document_widget.setVisible(False)


