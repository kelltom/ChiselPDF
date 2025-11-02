import os
import sys
from pathlib import Path

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QLineEdit, 
                             QFileDialog, QMessageBox, QGroupBox)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

try:
    from PyPDF2 import PdfReader, PdfWriter
except ImportError:
    print("PyPDF2 library is required. Installing...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "PyPDF2"])
    from PyPDF2 import PdfReader, PdfWriter


def parse_page_ranges(page_input):
    """
    Parse page ranges like '1-3,5,6-9,11' into a list of page numbers.
    Returns a sorted list of unique page numbers (1-indexed).
    """
    pages = set()
    parts = page_input.replace(" ", "").split(",")
    
    for part in parts:
        if not part:  # Skip empty parts
            continue
        if "-" in part:
            # Handle range like "1-3"
            start, end = part.split("-", 1)
            try:
                start_page = int(start)
                end_page = int(end)
                if start_page > end_page:
                    raise ValueError(f"Invalid range {part} (start > end)")
                pages.update(range(start_page, end_page + 1))
            except ValueError as e:
                raise ValueError(f"Invalid range format '{part}': {e}")
        else:
            # Handle single page number
            try:
                pages.add(int(part))
            except ValueError:
                raise ValueError(f"Invalid page number '{part}'")
    
    return sorted(pages)


def trim_pdf(input_path, page_numbers, output_path):
    """
    Create a new PDF with only the specified pages.
    
    Args:
        input_path: Path to the input PDF file
        page_numbers: List of page numbers to include (1-indexed)
        output_path: Path where the output PDF will be saved
    
    Returns:
        tuple: (success: bool, message: str, valid_pages: list)
    """
    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()
        
        total_pages = len(reader.pages)
        
        # Validate and add pages
        valid_pages = []
        invalid_pages = []
        for page_num in page_numbers:
            if 1 <= page_num <= total_pages:
                # Convert to 0-indexed for PyPDF2
                writer.add_page(reader.pages[page_num - 1])
                valid_pages.append(page_num)
            else:
                invalid_pages.append(page_num)
        
        if not valid_pages:
            return False, "No valid pages to include in the output PDF.", []
        
        # Write the output PDF
        with open(output_path, "wb") as output_file:
            writer.write(output_file)
        
        message = f"Successfully created PDF with {len(valid_pages)} pages"
        if invalid_pages:
            message += f"\n\nSkipped invalid pages: {invalid_pages} (PDF has {total_pages} pages)"
        
        return True, message, valid_pages
        
    except Exception as e:
        return False, f"Error processing PDF: {str(e)}", []


class PDFPageSelectorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_path = None
        self.output_path = None
        self.total_pages = 0
        
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("PDF Page Selector")
        self.setMinimumSize(400, 500)
        self.resize(400, 500)
        
        # Set application style
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QGroupBox {
                border: 1px solid #555555;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
                font-weight: bold;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QLabel {
                color: #ffffff;
            }
            QLineEdit {
                background-color: #3c3c3c;
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 3px;
                padding: 5px;
            }
            QPushButton {
                background-color: #0d47a1;
                color: #ffffff;
                border: none;
                border-radius: 3px;
                padding: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1565c0;
            }
            QPushButton:pressed {
                background-color: #0a3d91;
            }
            QPushButton:disabled {
                background-color: #555555;
                color: #888888;
            }
        """)
        
        # Central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # Title
        title_label = QLabel("PDF Page Selector")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        main_layout.addSpacing(10)
        
        # Input PDF group
        input_group = QGroupBox("Input PDF")
        input_layout = QVBoxLayout()
        input_layout.setSpacing(8)
        
        input_button_layout = QHBoxLayout()
        self.input_label = QLabel("No file selected")
        self.input_label.setStyleSheet("color: #888888; padding: 5px;")
        self.input_label.setWordWrap(True)
        input_button_layout.addWidget(self.input_label, 1)
        
        input_browse_btn = QPushButton("Browse...")
        input_browse_btn.setFixedWidth(100)
        input_browse_btn.clicked.connect(self.browse_input)
        input_button_layout.addWidget(input_browse_btn)
        
        input_layout.addLayout(input_button_layout)
        
        self.page_info_label = QLabel("")
        self.page_info_label.setStyleSheet("color: #aaaaaa; font-size: 10pt;")
        input_layout.addWidget(self.page_info_label)
        
        input_group.setLayout(input_layout)
        main_layout.addWidget(input_group)
        
        # Page ranges group
        page_group = QGroupBox("Page Ranges")
        page_layout = QVBoxLayout()
        page_layout.setSpacing(8)
        
        self.page_entry = QLineEdit()
        self.page_entry.setPlaceholderText("e.g., 1-3,5,6-9,11")
        page_layout.addWidget(self.page_entry)
        
        help_label = QLabel("Examples: 1-3,5,6-9,11  or  1,3,5  or  10-20")
        help_label.setStyleSheet("color: #aaaaaa; font-size: 9pt;")
        page_layout.addWidget(help_label)
        
        page_group.setLayout(page_layout)
        main_layout.addWidget(page_group)
        
        # Output PDF group
        output_group = QGroupBox("Output PDF")
        output_layout = QVBoxLayout()
        output_layout.setSpacing(8)
        
        output_button_layout = QHBoxLayout()
        self.output_label = QLabel("No output location selected")
        self.output_label.setStyleSheet("color: #888888; padding: 5px;")
        self.output_label.setWordWrap(True)
        output_button_layout.addWidget(self.output_label, 1)
        
        output_browse_btn = QPushButton("Browse...")
        output_browse_btn.setFixedWidth(100)
        output_browse_btn.clicked.connect(self.browse_output)
        output_button_layout.addWidget(output_browse_btn)
        
        output_layout.addLayout(output_button_layout)
        output_group.setLayout(output_layout)
        main_layout.addWidget(output_group)
        
        # Process button
        self.process_button = QPushButton("Create Trimmed PDF")
        self.process_button.setMinimumHeight(40)
        process_font = QFont()
        process_font.setPointSize(11)
        process_font.setBold(True)
        self.process_button.setFont(process_font)
        self.process_button.clicked.connect(self.process_pdf)
        self.process_button.setEnabled(False)
        main_layout.addWidget(self.process_button)
        
        # Status label
        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.status_label)
        
        # Add stretch to push everything to the top
        main_layout.addStretch()
        
    def browse_input(self):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select Input PDF",
            "",
            "PDF files (*.pdf);;All files (*.*)"
        )
        
        if filename:
            self.input_path = filename
            # Truncate long paths for display
            display_name = filename if len(filename) < 80 else "..." + filename[-77:]
            self.input_label.setText(display_name)
            self.input_label.setStyleSheet("color: #ffffff; padding: 5px;")
            
            # Get page count
            try:
                reader = PdfReader(filename)
                self.total_pages = len(reader.pages)
                self.page_info_label.setText(f"Total pages: {self.total_pages}")
                
                # Suggest default output
                input_file = Path(filename)
                default_output = input_file.parent / f"{input_file.stem}_trimmed{input_file.suffix}"
                self.output_path = str(default_output)
                display_output = str(default_output) if len(str(default_output)) < 80 else "..." + str(default_output)[-77:]
                self.output_label.setText(display_output)
                self.output_label.setStyleSheet("color: #ffffff; padding: 5px;")
                
                self.update_process_button()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not read PDF: {str(e)}")
                self.input_path = None
                self.total_pages = 0
                self.input_label.setText("No file selected")
                self.input_label.setStyleSheet("color: #888888; padding: 5px;")
                self.page_info_label.setText("")
    
    def browse_output(self):
        if self.input_path:
            input_file = Path(self.input_path)
            initial_file = f"{input_file.stem}_trimmed{input_file.suffix}"
            initial_dir = str(input_file.parent)
        else:
            initial_file = "output_trimmed.pdf"
            initial_dir = str(Path.home())
        
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Save Trimmed PDF As",
            f"{initial_dir}/{initial_file}",
            "PDF files (*.pdf);;All files (*.*)"
        )
        
        if filename:
            self.output_path = filename
            display_name = filename if len(filename) < 80 else "..." + filename[-77:]
            self.output_label.setText(display_name)
            self.output_label.setStyleSheet("color: #ffffff; padding: 5px;")
            self.update_process_button()
    
    def update_process_button(self):
        if self.input_path and self.output_path:
            self.process_button.setEnabled(True)
        else:
            self.process_button.setEnabled(False)
    
    def process_pdf(self):
        # Validate page input
        page_input = self.page_entry.text().strip()
        if not page_input:
            QMessageBox.warning(self, "Error", "Please enter page ranges.")
            return
        
        try:
            page_numbers = parse_page_ranges(page_input)
        except ValueError as e:
            QMessageBox.warning(self, "Error", f"Invalid page format:\n{str(e)}")
            return
        
        if not page_numbers:
            QMessageBox.warning(self, "Error", "No valid page numbers found.")
            return
        
        # Check if output file exists
        if self.output_path and os.path.exists(self.output_path):
            reply = QMessageBox.question(
                self,
                "Confirm Overwrite",
                f"File already exists:\n{self.output_path}\n\nOverwrite?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
        
        # Process the PDF
        self.status_label.setText("Processing...")
        self.status_label.setStyleSheet("color: #aaaaaa;")
        QApplication.processEvents()  # Update UI
        
        success, message, valid_pages = trim_pdf(self.input_path, page_numbers, self.output_path)
        
        if success:
            self.status_label.setText(f"Success! Saved to: {self.output_path}")
            self.status_label.setStyleSheet("color: #4caf50;")
            QMessageBox.information(self, "Success", message)
        else:
            self.status_label.setText("Failed")
            self.status_label.setStyleSheet("color: #f44336;")
            QMessageBox.critical(self, "Error", message)


def main():
    app = QApplication(sys.argv)
    window = PDFPageSelectorApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
