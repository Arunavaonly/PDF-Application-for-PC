from ui import Ui_MainWindow
from PySide6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QLineEdit, QWidget, QPushButton, QListWidgetItem, QLabel, QHBoxLayout
from PySide6.QtCore import Qt,QTimer, QSize
from PySide6.QtGui import QMovie, QIcon

import os
import fitz


class EnigmaPdf(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("EnigmaPDF")

        # Progress bar setup
        self.progress_bar_movie = QMovie("progress_bar.gif")
        self.progress_bar_label.setMovie(self.progress_bar_movie)
        #self.progress_bar_label.setVisible(False)  # Hide by default
        self.progress_update_label.hide()

        # Set a custom font size for the status bar messages
        font = self.statusBar().font()
        font.setPointSize(10)  # Adjust the size as needed
        self.statusBar().setFont(font)

        # Optionally, set a stylesheet to make the text bold and in italic
        self.statusBar().setStyleSheet("QStatusBar { font-weight: bold; font-style: italic; }")


        self.mainremove_button.clicked.connect(self.switch_removepage)
        self.mainconvert_button.clicked.connect(self.switch_convertpage)
        self.mainextract_button.clicked.connect(self.switch_extractpage)
        self.mainsplit_button.clicked.connect(self.switch_splitpage)
        self.maininsert_button.clicked.connect(self.switch_insertpage)
        self.mainmerge_button.clicked.connect(self.switch_mergepage)
        self.mainpassword_button.clicked.connect(self.switch_protectpage)
        self.mainwatermark_button.clicked.connect(self.switch_watermarkpage)
        self.mainresize_button.clicked.connect(self.switch_resizepage)

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

        self.resize_button.clicked.connect(lambda: self.blink_and_execute(self.resize_button, self.reduce_size))
        self.closed_eye.clicked.connect(self.toggle_password_visibility)

        self.convert_from_dropdown.currentIndexChanged.connect(self.update_convert_to)
        
        self.selected_file_path = ""
        self.selected_output_path = ""
        self.document_for_insert = ""
        self.file_list = []
        self.is_password_visible = False
        self.select_document_widget.setVisible(False)

        # Ensure convertPage is shown and mainConvert is checked
        self.mainconvert_button.setChecked(True)
        self.stackedWidget.setCurrentWidget(self.convert_page)
    

    def blink_button(self, button):
        # Hide the button
        button.setVisible(False)
        # Create a QTimer to show the button after a few milliseconds
        QTimer.singleShot(250, lambda: button.setVisible(True))

    def show_message(self, message, duration=2000):
        # Show the message
        self.progress_update_label.setText(message)
        self.progress_update_label.show()

        # Use QTimer.singleShot to hide the message after the specified duration
        QTimer.singleShot(duration, self.hide_message)

    def hide_message(self):
        self.progress_update_label.hide()

    def show_progress_bar(self):
        """Show and start the progress_bar."""
        self.progress_bar_label.setVisible(True)
        self.progress_bar_movie.start()
        QApplication.processEvents()  # Ensure the UI updates immediately

    def hide_progress_bar(self):
        """Hide and stop the progress_bar."""
        self.progress_bar_label.setVisible(False)
        self.progress_bar_movie.stop()

    def blink_and_execute(self, button, func):
        self.blink_button(button)  # Make the button blink
 
        #self.set_all_enabled(False)  # Disable all buttons during task execution
        QApplication.setOverrideCursor(Qt.WaitCursor)  # Change cursor to busy

        # Start the function with a timer to ensure it runs for 2 seconds
        def run_task_with_waitcursor():
            try:
                func()  # Execute the task
            finally:
                #self.set_all_enabled(True)  # Re-enable all buttons
                QApplication.restoreOverrideCursor()  # Restore default cursor

        QTimer.singleShot(300, run_task_with_waitcursor)  # Start immediately

        
    def set_all_enabled(self, enabled):
        self.remove_button.setEnabled(enabled)
        self.convert_button.setEnabled(enabled)
        self.extract_button.setEnabled(enabled)
        self.insert_button.setEnabled(enabled)
        self.split_button.setEnabled(enabled)
        self.watermark_button.setEnabled(enabled)
        self.merge_button.setEnabled(enabled)
        self.password_button.setEnabled(enabled)
        self.select_button.setEnabled(enabled)
        self.output_button.setEnabled(enabled)
        self.selectfolder_button.setEnabled(enabled)
        self.addfiles_button.setEnabled(enabled)


    def get_file_filter(self):
        # Check which button is checked and return the appropriate file filter
        if self.mainconvert_button.isChecked():
            if self.convert_from_dropdown.currentText() == "Image":
                return "Image Files (*.png *.jpg *.jpeg )"
            else:
                return "PDF Files (*.pdf)"
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
            if self.mainconvert_button.isChecked() and self.convert_from_dropdown.currentText() == "Image":  # Modified to add files based on conversion option (Lines 17-23)
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
        self.show_files.clear()  # Clear existing list
        for index, file in enumerate(self.file_list):
            # Create a custom QWidget to hold the file label and remove button
            file_widget = QWidget()
            file_layout = QHBoxLayout(file_widget)
            
            # Create a QLabel to display the file name
            file_label = QLabel(f"{index + 1}. {file}")
            
            # Create a QPushButton with an "X" icon to remove the file
            remove_button = QPushButton()
            remove_button.setIcon(QIcon("cross_icon.png"))  # Replace with the actual path to your icon
            remove_button.setFixedSize(QSize(15, 15))  # Set the size of the remove button
            remove_button.setFlat(True)
            
            # Connect the remove button to a function to remove the file
            remove_button.clicked.connect(lambda _, f=file: self.remove_file(f))
            
            # Add the QLabel and QPushButton to the layout
            file_layout.addWidget(remove_button)
            #file_layout.addStretch()  # Optional: adds spacing to the right of the button
            file_layout.addWidget(file_label)
            
            # Create a QListWidgetItem and set the custom QWidget as its item
            list_item = QListWidgetItem(self.show_files)
            list_item.setSizeHint(file_widget.sizeHint())  # Set size of list item
            self.show_files.addItem(list_item)
            self.show_files.setItemWidget(list_item, file_widget)

    def remove_file(self, file_to_remove):
        """Removes the selected file from the file list and updates the UI."""
        if file_to_remove in self.file_list:
            self.file_list.remove(file_to_remove)
            self.add_file()  # Refresh the list to reflect the removed file

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
        self.set_all_enabled(False)
        try:
            self.show_progress_bar()

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
            self.show_message("Success: Insert action completed successfully.", duration =500)
            QApplication.processEvents()
            self.statusBar().showMessage("inserted.pdf saved", 2000)

        except Exception as e:
                self.hide_progress_bar()
                QMessageBox.critical(self, "Error", f"An error occured while inserting: {e}")
        
        self.file_path.clear()
        self.insert_document_path.clear()
        self.selected_file_path = ""
        self.document_for_insert = ""
        self.page_range_insert.setText("Page Range: ")
        self.insert_at_dropdown.clear()
        self.start_page_dropdown.clear()
        self.end_page_dropdown.clear()
        QTimer.singleShot(1000, self.hide_progress_bar)  # Ensure progress_bar runs for at least 1 seconds

        self.set_all_enabled(True)


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
        try:
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
                    self.show_message(f"Success: Text extracted from {filename} successfully.", duration = 500)
                    QApplication.processEvents()

                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to extract text: {e}")
                    continue
            self.statusBar().showMessage("extracted_text.txt saved", 2000)
            QApplication.processEvents()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to extract text: {e}")



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
                self.show_message(f"Success: Images extracted from {filename} successfully.", duration =500)
                QApplication.processEvents()
                self.statusBar().showMessage("All images extracted successfully", 2000)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to extract images from {filename}: {e}")
                continue

    def extract(self):
        # Get the selected option from the dropdown        
        if len(self.file_list) ==0:
            QMessageBox.warning(self, "Input Error", "Please select a PDF file.")
            return
        
        self.set_all_enabled(False)
        if self.selected_output_path:
            self.show_progress_bar()  # Show progress_bar
            if self.extract_dropdown.currentText() == "Text":
                self.text_extract()

            elif self.extract_dropdown.currentText() == "Images":
                self.image_extract()
            else:
                QMessageBox.warning(self, "Input Error", "Please select a valid extraction option.")
                self.set_all_enabled(True)
                return
        else:
            QMessageBox.warning(self,  "No Output Path", "Please select an output path.")
            self.set_all_enabled(True)
            return
        QTimer.singleShot(1000, self.hide_progress_bar)  # Ensure progress_bar runs for at least 1 seconds

        self.show_files.clear()
        self.file_path.clear()
        self.folder_path.clear()

        self.file_list = []
        self.selected_file_path = ''
        self.set_all_enabled(True)

        
    def split_pdf(self):        
        input_pdf = self.selected_file_path
        if input_pdf == '':
            QMessageBox.warning(self, "Input Error", "Please select the input PDF.")
            return

        self.set_all_enabled(False)
        input_pdf = self.selected_file_path
        try:
            self.show_progress_bar()  # Show progress_bar
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

            self.show_message("Success: PDF splitted into two parts successfully.", duration =1000)
            self.statusBar().showMessage(f"{self.output1.text()}.pdf and {self.output2.text()}.pdf saved", 2000)
            QApplication.processEvents()

        except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to split the PDF: {e}")

        self.selected_file_path = ""
        self.file_path.clear()
        self.page_dropdown.clear()
        self.page_range.setText("Page Range: ")
        QTimer.singleShot(1000, self.hide_progress_bar)  # Ensure progress_bar runs for at least 1 seconds

        self.set_all_enabled(True)



    def show_page_range(self):
        if self.mainsplit_button.isChecked() and self.selected_file_path:
            pdf_file = self.selected_file_path
            document = fitz.open(pdf_file)
            pages = document.page_count
            self.page_range.setText(f"Page Range: 1-{pages}")
            self.page_dropdown.clear()
            for page_number in range(1, pages+1):
                self.page_dropdown.addItem(str(page_number))

    def toggle_password_visibility(self):
        if self.is_password_visible:
            self.password_text.setEchoMode(QLineEdit.Password)  # hide password
            self.closed_eye.setIcon(QIcon('closed_eye.png'))
        else:
            self.password_text.setEchoMode(QLineEdit.Normal)  # Show password
            self.closed_eye.setIcon(QIcon('opened_eye.jpg'))

        # Toggle the visibility flag
        self.is_password_visible = not self.is_password_visible

    def encrypt_pdf(self):
        password = self.password_text.text()
        output_path = self.selected_output_path
        # Validate the inputs
        if len(self.file_list) == 0 or not output_path or not password:
            QMessageBox.warning(self, "Input Error", "Please provide all required inputs.")
            return
        
        self.set_all_enabled(False)
        try:
            self.show_progress_bar()  # Show progress_bar

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
                    self.show_message(f'encrypted_{base_name} saved', duration = 500)

                    QApplication.processEvents()
                
                except Exception as e:
                    self.hide_progress_bar()
                    QMessageBox.critical(self, "Error", f"Failed to encrypt {base_name}: {e}")
                    continue
            QTimer.singleShot(1000, self.hide_progress_bar)  # Ensure progress_bar runs for at least 1 seconds
            self.statusBar().showMessage("Success: PDFs encrypted successfully.", 2000)
            QApplication.processEvents()


        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to encrypt PDF: {e}")

        # Clear the paths and file list
        self.selected_file_path = ""
        self.file_list = []

        self.show_files.clear()
        self.file_path.clear()
        self.folder_path.clear()
        self.set_all_enabled(True)
        

    def watermark(self):
        watermark = self.watermark_text.text()
        output_path = self.selected_output_path
        if len(self.file_list) == 0 or not output_path or not watermark:
            QMessageBox.warning(self, "Input Error", "Please provide all required inputs.")
            return
        self.set_all_enabled(False)

        for file in self.file_list:
            try:
                self.show_progress_bar()  # Show progress_bar

                # Open the PDF file
                pdf = fitz.open(file)  # Replace with your PDF file path

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
                    # Construct the output file path
                    base_name = os.path.basename(file)
                    output_file_path = os.path.join(output_path, f"watermarked_{base_name}")

                # Save the watermarked PDF
                pdf.save(output_file_path)  # Replace with your desired output PDF file name

                # Close the PDF file
                pdf.close()
                self.show_message(f"Success:{base_name} watermarked successfully.'watermarked_{base_name}' saved", duration=500)
                QApplication.processEvents()

            except Exception as e:
                QMessageBox.critical(self, "Error",f"Failed to watermark {base_name}: {e}")
                continue
            
        QTimer.singleShot(1000, self.hide_progress_bar)  # Ensure progress_bar runs for at least 1 seconds
        self.statusBar().showMessage("All files watermarked", 2000)
        QApplication.processEvents()


        self.show_files.clear()
        self.file_path.clear()
        self.folder_path.clear()

        self.file_list = []
        self.selected_file_path = ''
        self.set_all_enabled(True)


    def merge_pdf(self):
        if len(self.file_list) ==0 or self.selected_output_path == "":
            QMessageBox.warning(self, "Input Error", "Please provide all required inputs.")
            return
        if len(self.file_list) ==1:
            QMessageBox.warning(self, "Single File Error", "Please provide at least two PDFs to start merging")
            return
        self.set_all_enabled(False)
        try:
            self.show_progress_bar()

            merged = fitz.open()
            for file in self.file_list:
                pdf_document = fitz.open(file)
                merged.insert_pdf(pdf_document)
                pdf_document.close()
            merged_pdf = os.path.join(self.selected_output_path, "merged_output.pdf").replace("\\", '/')
            merged.save(merged_pdf)
            merged.close()
            self.show_message("Success: PDFs merged successfully.'merged_output.pdf' saved", duration =2000)
            QApplication.processEvents()
        except Exception as e:
            QMessageBox.critical(self, "Error","Failed to merge PDFs: {e}")
        
        QTimer.singleShot(1000, self.hide_progress_bar)  # Ensure progress_bar runs for at least 1 seconds

            
        self.show_files.clear()
        self.file_path.clear()
        self.folder_path.clear()

        self.file_list = []
        self.selected_file_path = ''

        self.set_all_enabled(True)
    

    def convert(self):
        if len(self.file_list) ==0 or self.selected_output_path == "":
            QMessageBox.warning(self, "Input Error", "Please select all required fields.")
            return
        self.set_all_enabled(False)

        try:
            self.show_progress_bar()
            if self.convert_from_dropdown.currentText() == "Image" and self.convert_to_dropdown.currentText()== "PDF":
                self.img2pdf()
            elif self.convert_to_dropdown.currentText() == "Image(.jpg)":
                self.pdf2img("jpg")
            elif self.convert_to_dropdown.currentText() == "Image(.jpeg)":
                self.pdf2img("jpeg")
            elif self.convert_to_dropdown.currentText() == "Image(.png)":
                self.pdf2img("png")
            else:
                QMessageBox.warning(self, "Input Error", "Please select a valid conversion option.")
                self.set_all_enabled(True)
                self.hide_progress_bar()
                return
            
        except Exception as e:
            self.hide_progress_bar()
            self.set_all_enabled(True)
            QMessageBox.critical(self, "Error",f"Error: {e}")
            return        
        QTimer.singleShot(1000, self.hide_progress_bar)  # Ensure progress_bar runs for at least 1 seconds
        QApplication.processEvents()

        self.show_files.clear()
        self.file_path.clear()
        self.folder_path.clear()

        self.file_list = []
        self.selected_file_path = ''

        self.set_all_enabled(True)


    def update_convert_to(self):
        self.convert_to_dropdown.clear()

        if self.convert_from_dropdown.currentText() == 'PDF':
            self.convert_to_dropdown.addItems(['Image(.jpg)', 'Image(.jpeg)', 'Image(.png)'])
        else:
            self.convert_to_dropdown.addItems(['PDF'])



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
                self.show_message(f"Success: pages saved as images from {filename} successfully.", duration = 500)
                QApplication.processEvents()
                pdf.close()  # Close the PDF
            except Exception as e:
                QMessageBox.critical(self, f"Error","Error: {e} happened in {filename}")
                continue
        self.statusBar().showMessage("All images saved", 2000)



    def img2pdf(self):
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
                    self.show_message(f"Success: {filename} converted into PDF successfully.", duration = 500)
                    QApplication.processEvents()
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Error: {e} in {filename}")
                    continue
            output_pdf = os.path.join(self.selected_output_path, "output.pdf")     
            doc.save(output_pdf)  # Save the new PDF
            doc.close()
            self.statusBar().showMessage("output.pdf saved", 2000)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error: {e}")

    def remove(self):
        input_file = self.selected_file_path
        output_path = self.selected_output_path
        if not input_file or not output_path:
            QMessageBox.warning(self, "Input", "Please provide all required fields.")
            return
        
        self.set_all_enabled(False)
        try:
            self.show_progress_bar()
            file_handle = fitz.open(input_file)
            start = int(self.start_page_dropdown_remove.currentText()) #Start Index
            end = int(self.end_page_dropdown_remove.currentText())    # End index
            if not end - start < file_handle.page_count -1:
                QMessageBox.warning(self,"Issue with Page Range", "Please change the page range")
            else:
                file_handle.delete_pages(start -1,end -1)
                filename = os.path.basename(input_file)
                output_file = os.path.join(output_path, f"pages_deleted_{filename}")

                file_handle.save(output_file)
                self.show_message("Success: pages removed", duration = 2000)
                self.statusBar().showMessage(f"pages_deleted_{filename} saved", 2000)

        except Exception as e:
                QMessageBox.critical(self, "Error", f'An error occured {e}')

        QTimer.singleShot(1000, self.hide_progress_bar)  # Ensure progress_bar runs for at least 1 seconds
        QApplication.processEvents()

        self.selected_file_path = ""

        self.start_page_dropdown_remove.clear()
        self.end_page_dropdown_remove.clear()
        self.page_range_remove.setText("Page Range")
        self.file_path.clear()
        self.set_all_enabled(True)
    
    def reduce_size(self):
        output_path = self.selected_output_path
        if len(self.file_list) == 0 or not output_path:
            QMessageBox.warning(self, "Input", "Please provide all required fields.")
            return
        self.set_all_enabled(False)
        for file in self.file_list:
            try:
                self.show_progress_bar()
                base_name = os.path.basename(file)
                doc = fitz.Document(file)
                doc_new = fitz.Document()
                for page in doc:
                    pixmap = page.get_pixmap(colorspace= fitz.csRGB, dpi= int(self.dpi_dropdown.currentText()), annots=False)
                    new_page = doc_new.new_page(-1)
                    xref = new_page.insert_image(rect=new_page.bound(), pixmap=pixmap)

                output_pdf_path = os.path.join(output_path, f'resized_{base_name}')
                doc_new.save(output_pdf_path, garbage=4, deflate=True, deflate_images=True, deflate_fonts=False, pretty=True)
                doc.close()
                doc_new.close()
                self.show_message(f'{base_name} successfully resized', 500)
                QApplication.processEvents()

            except Exception as e:
                QMessageBox.critical(self, "Error", f'An error occured in {base_name}: {e}')
                continue
        self.statusBar().showMessage("All files are resized", 2000)
        QTimer.singleShot(1000, self.hide_progress_bar)  # Ensure progress_bar runs for at least 1 seconds
        QApplication.processEvents()

        self.show_files.clear()
        self.file_path.clear()
        self.folder_path.clear()

        self.file_list = []
        self.selected_file_path = ''

        self.set_all_enabled(True)


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
        self.select_document_widget.setVisible(False)
        self.selectfolder_button.setEnabled(False)
        self.addfiles_button.setEnabled(False)


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
        self.select_document_widget.setVisible(False)
        self.selectfolder_button.setEnabled(True)
        self.addfiles_button.setEnabled(True)


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
        self.select_document_widget.setVisible(False)
        self.selectfolder_button.setEnabled(True)
        self.addfiles_button.setEnabled(True)


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
        self.select_document_widget.setVisible(False)
        self.selectfolder_button.setEnabled(True)
        self.addfiles_button.setEnabled(True)


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
        self.select_document_widget.setVisible(False)
        self.selectfolder_button.setEnabled(True)
        self.addfiles_button.setEnabled(True)


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
        self.select_document_widget.setVisible(False)
        self.selectfolder_button.setEnabled(True)
        self.addfiles_button.setEnabled(True)

    def switch_resizepage(self):
        self.folder_path.clear()
        self.show_files.clear()
        self.file_path.clear()
        self.output_path.clear()
        self.page_dropdown.clear()
        self.page_range.setText("Page Range: ")

        self.selected_file_path = ""
        self.selected_output_path = ""
        self.file_list = []

        self.stackedWidget.setCurrentWidget(self.resize_page)
        self.select_document_widget.setVisible(False)
        self.selectfolder_button.setEnabled(True)
        self.addfiles_button.setEnabled(True)





