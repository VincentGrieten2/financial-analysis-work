import os
import sys
from pathlib import Path
from typing import List, Callable
import subprocess
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QListWidget, QProgressBar,
    QFrame, QMessageBox, QScrollArea
)
from PySide6.QtCore import Qt, QMimeData
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont

class DropZone(QFrame):
    def __init__(self, accept_type: str, on_drop, parent=None):
        super().__init__(parent)
        self.accept_type = accept_type
        self.on_drop = on_drop
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        self.setStyleSheet("""
            DropZone {
                background-color: #F5F5F5;
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                padding: 8px;
                min-height: 40px;
            }
            DropZone:hover {
                background-color: #EEEEEE;
                border: 1px solid #BDBDBD;
            }
        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if all(url.toLocalFile().lower().endswith(self.accept_type) for url in urls):
                event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        files = [url.toLocalFile() for url in urls]
        self.on_drop(files)

class FinancialAnalysisGUI(QMainWindow):
    def __init__(self, process_callback: Callable[[str, List[str]], bool]):
        super().__init__()
        self.process_callback = process_callback
        self.pdf_files: List[str] = []
        self.template_file: str = ""
        self.output_file: str = ""
        
        self.setWindowTitle("Financial Analysis Tool")
        self.setMinimumSize(650, 650)  # Increased minimum size
        self.setup_ui()

    def setup_ui(self):
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(4)  # Reduced from 8
        main_layout.setContentsMargins(6, 6, 6, 6)  # Reduced from 8,8,8,8

        # Create a scroll area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # Create a container widget for the scroll area
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setSpacing(4)  # Reduced from 8
        scroll_layout.setContentsMargins(2, 2, 2, 2)  # Reduced from 4,4,4,4

        # Template Section
        template_frame = QFrame()
        template_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 6px;
                padding: 6px;
            }
        """)
        template_layout = QVBoxLayout(template_frame)
        template_layout.setSpacing(4)  # Reduced from 8
        
        # Template Header
        template_header = QLabel("Excel Template")
        template_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        template_layout.addWidget(template_header)

        # Template Buttons
        template_btn_layout = QHBoxLayout()
        self.select_template_btn = QPushButton("Select Template")
        self.select_template_btn.clicked.connect(self._select_template)
        self.clear_template_btn = QPushButton("Clear Template")
        self.clear_template_btn.clicked.connect(self._clear_template)
        self.clear_template_btn.setEnabled(False)
        
        template_btn_layout.addWidget(self.select_template_btn)
        template_btn_layout.addStretch()
        template_btn_layout.addWidget(self.clear_template_btn)
        template_layout.addLayout(template_btn_layout)

        # Template Drop Zone
        template_drop_container = QWidget()
        template_drop_layout = QVBoxLayout(template_drop_container)
        self.template_label = QLabel("No template selected")
        self.template_label.setFont(QFont("Segoe UI", 10))
        self.template_label.setWordWrap(True)  # Enable word wrap
        template_drop_hint = QLabel("Drag and drop Excel template here")
        template_drop_hint.setStyleSheet("color: #757575; font-style: italic;")
        
        template_drop_layout.addWidget(self.template_label)
        template_drop_layout.addWidget(template_drop_hint)
        
        self.template_drop_zone = DropZone(".xlsx", self._on_template_drop)
        self.template_drop_zone.setLayout(template_drop_layout)
        template_layout.addWidget(self.template_drop_zone)
        
        scroll_layout.addWidget(template_frame)

        # PDF Section
        pdf_frame = QFrame()
        pdf_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 6px;
                padding: 6px;
            }
        """)
        pdf_layout = QVBoxLayout(pdf_frame)
        pdf_layout.setSpacing(4)  # Reduced from 8
        
        # PDF Header
        pdf_header = QLabel("PDF Files (Max 5)")
        pdf_header.setFont(QFont("Segoe UI", 12, QFont.Bold))
        pdf_layout.addWidget(pdf_header)

        # PDF Buttons
        pdf_btn_layout = QHBoxLayout()
        self.select_pdf_btn = QPushButton("Select PDFs")
        self.select_pdf_btn.clicked.connect(self._select_pdfs)
        self.clear_pdfs_btn = QPushButton("Clear All PDFs")
        self.clear_pdfs_btn.clicked.connect(self._clear_pdfs)
        self.clear_pdfs_btn.setEnabled(False)
        
        pdf_btn_layout.addWidget(self.select_pdf_btn)
        pdf_btn_layout.addStretch()
        pdf_btn_layout.addWidget(self.clear_pdfs_btn)
        pdf_layout.addLayout(pdf_btn_layout)

        # PDF List and Drop Zone
        pdf_drop_container = QWidget()
        pdf_drop_layout = QVBoxLayout(pdf_drop_container)
        
        self.pdf_list = QListWidget()
        self.pdf_list.setStyleSheet("""
            QListWidget {
                background-color: white;
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                min-height: 60px;
                max-height: 100px;
                padding: 2px;
            }
        """)
        self.pdf_list.setFont(QFont("Segoe UI", 10))
        self.pdf_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.pdf_list.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        pdf_drop_layout.addWidget(self.pdf_list)
        
        pdf_drop_hint = QLabel("Drag and drop PDF files here")
        pdf_drop_hint.setStyleSheet("color: #757575; font-style: italic;")
        pdf_drop_layout.addWidget(pdf_drop_hint)
        
        self.pdf_drop_zone = DropZone(".pdf", self._on_pdf_drop)
        self.pdf_drop_zone.setLayout(pdf_drop_layout)
        pdf_layout.addWidget(self.pdf_drop_zone)
        
        scroll_layout.addWidget(pdf_frame)

        # Status Section
        status_frame = QFrame()
        status_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 6px;
                padding: 6px;
                margin-top: 0px;
            }
        """)
        status_layout = QVBoxLayout(status_frame)
        status_layout.setSpacing(2)  # Reduced from 4
        status_layout.setContentsMargins(6, 2, 6, 6)  # Reduced from 8,4,8,8
        
        # Status Label
        self.status_label = QLabel()
        self.status_label.setFont(QFont("Segoe UI", 10))
        self.status_label.setWordWrap(True)
        status_layout.addWidget(self.status_label)
        
        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                text-align: center;
                margin-top: 0px;
                margin-bottom: 4px;
            }
            QProgressBar::chunk {
                background-color: #1976D2;
            }
        """)
        self.progress_bar.hide()
        status_layout.addWidget(self.progress_bar)
        
        # Process and Open Buttons
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(0, 0, 0, 0)  # Remove button layout margins
        self.process_button = QPushButton("Process Files")
        self.process_button.clicked.connect(self._process_files)
        self.process_button.setEnabled(False)
        self.process_button.setStyleSheet("""
            QPushButton {
                background-color: #1976D2;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1565C0;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """)
        
        self.open_output_button = QPushButton("Open Output File")
        self.open_output_button.clicked.connect(self._open_output_file)
        self.open_output_button.setEnabled(False)
        self.open_output_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #388E3C;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """)
        
        button_layout.addWidget(self.process_button)
        button_layout.addStretch()
        button_layout.addWidget(self.open_output_button)
        status_layout.addLayout(button_layout)
        
        scroll_layout.addWidget(status_frame)

        # Set the scroll widget as the scroll area's widget
        scroll_area.setWidget(scroll_widget)
        main_layout.addWidget(scroll_area)

    def _on_template_drop(self, files):
        """Handle template file drop"""
        if files:
            file_path = files[0]  # Take only the first file
            if file_path.lower().endswith('.xlsx'):
                self._set_template(file_path)
            else:
                QMessageBox.warning(self, "Invalid File", "Please drop an Excel (.xlsx) file")

    def _on_pdf_drop(self, files):
        """Handle PDF file drop"""
        for file_path in files:
            if file_path.lower().endswith('.pdf'):
                self._add_pdf(file_path)
            else:
                QMessageBox.warning(
                    self, 
                    "Invalid File", 
                    f"Skipped {Path(file_path).name} - not a PDF file"
                )

    def _select_template(self):
        """Open file dialog to select template"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel Template",
            "",
            "Excel files (*.xlsx)"
        )
        if file_path:
            self._set_template(file_path)

    def _select_pdfs(self):
        """Open file dialog to select PDFs"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PDF Files",
            "",
            "PDF files (*.pdf)"
        )
        for file in files:
            self._add_pdf(file)

    def _set_template(self, path: str):
        """Set the template file"""
        print(f"Setting template file: {path}")  # Debug print
        self.template_file = path
        self.template_label.setText(f"Template: {Path(path).name}")
        self.clear_template_btn.setEnabled(True)
        self._update_process_button()
        print(f"Current template file: {self.template_file}")  # Debug print

    def _clear_template(self):
        """Clear the template selection"""
        self.template_file = ""
        self.template_label.setText("No template selected")
        self.clear_template_btn.setEnabled(False)
        self._update_process_button()
        # Reset output file when clearing template
        self.output_file = ""
        self.open_output_button.setEnabled(False)

    def _add_pdf(self, path: str):
        """Add a PDF file to the list"""
        print(f"Adding PDF file: {path}")  # Debug print
        if len(self.pdf_files) >= 5:
            QMessageBox.warning(self, "Maximum Files", "Maximum 5 PDF files allowed!")
            return
        
        if path not in self.pdf_files:
            self.pdf_files.append(path)
            self.pdf_list.addItem(Path(path).name)
            self.clear_pdfs_btn.setEnabled(True)
            self._update_process_button()
            print(f"Current PDF files: {self.pdf_files}")  # Debug print

    def _clear_pdfs(self):
        """Clear all PDF files"""
        self.pdf_files.clear()
        self.pdf_list.clear()
        self.clear_pdfs_btn.setEnabled(False)
        self._update_process_button()
        # Reset output file when clearing PDFs
        self.output_file = ""
        self.open_output_button.setEnabled(False)

    def _update_process_button(self):
        """Update the state of the process button"""
        self.process_button.setEnabled(bool(self.template_file and self.pdf_files))

    def _process_files(self):
        """Process the selected files"""
        try:
            print(f"Processing files - Template: {self.template_file}, PDFs: {self.pdf_files}")  # Debug print
            # Disable process button and show progress
            self.process_button.setEnabled(False)
            self.status_label.setText("Processing files...")
            self.status_label.setStyleSheet("color: #1976D2;")
            self.progress_bar.setRange(0, 0)  # Indeterminate mode
            self.progress_bar.show()
            QApplication.processEvents()
            
            success = self.process_callback(self.template_file, self.pdf_files)
            
            if success:
                # Find the most recent financial_analysis_*.xlsx file
                output_files = [f for f in os.listdir('.') if f.startswith('financial_analysis_') and f.endswith('.xlsx')]
                if output_files:
                    # Sort by modification time, most recent first
                    output_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                    self.output_file = output_files[0]
                    self.open_output_button.setEnabled(True)
                    self.status_label.setText("Processing completed successfully!")
                    self.status_label.setStyleSheet("color: #4CAF50;")
                    QMessageBox.information(
                        self,
                        "Success",
                        "Financial analysis completed successfully!"
                    )
                else:
                    self.status_label.setText("Error: Output file not found")
                    self.status_label.setStyleSheet("color: #D32F2F;")
                    QMessageBox.critical(self, "Error", "Output file not found!")
            else:
                self.status_label.setText(
                    "Error during processing. Check parser.log for details."
                )
                self.status_label.setStyleSheet("color: #D32F2F;")
                QMessageBox.critical(
                    self,
                    "Error",
                    "An error occurred during processing. Check parser.log for details."
                )
        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")
            self.status_label.setStyleSheet("color: #D32F2F;")
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
        finally:
            # Hide progress bar and restore process button
            self.progress_bar.hide()
            self.process_button.setEnabled(True)

    def _open_output_file(self):
        """Open the output Excel file"""
        if not self.output_file or not os.path.exists(self.output_file):
            QMessageBox.critical(self, "Error", "Output file not found!")
            return
        
        try:
            if os.name == 'nt':  # Windows
                os.startfile(self.output_file)
            else:  # macOS and Linux
                subprocess.run(['xdg-open' if os.name == 'posix' else 'open', self.output_file])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open file: {str(e)}")

def run_gui(process_callback: Callable[[str, List[str]], bool]):
    """Run the GUI application"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Use Fusion style for a modern look
    window = FinancialAnalysisGUI(process_callback)
    window.show()
    sys.exit(app.exec()) 