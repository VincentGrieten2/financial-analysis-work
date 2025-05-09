import os
import sys
from pathlib import Path
from typing import List, Callable, Tuple
import subprocess
<<<<<<< HEAD
import win32gui
import win32con
import win32api
from ctypes import windll, create_unicode_buffer, c_uint
from ctypes.wintypes import DWORD

class FinancialAnalysisGUI:
    def __init__(self, process_callback: Callable[[str, List[str]], str]):
        """
        Initialize the GUI.
        :param process_callback: Callback function that takes (template_path, pdf_paths) and returns output_path or None on failure
        """
=======
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
>>>>>>> new-branch-name
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
        
<<<<<<< HEAD
        # Configure grid weights
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(2, weight=1)
        
        # Enable OLE drag and drop
        windll.shell32.DragAcceptFiles(self.root.winfo_id(), True)
        
        self._create_widgets()
        self._setup_drag_drop()
        
    def _create_widgets(self):
        """Create all GUI widgets"""
        # Template Section
        template_frame = ttk.LabelFrame(self.root, text="Excel Template (Drag & Drop Enabled)", padding=10)
        template_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        template_frame.grid_columnconfigure(0, weight=1)
        
        self.template_label = ttk.Label(
            template_frame, 
            text="Drag & drop or click 'Select Template' to choose an Excel file",
            background='white',
            relief='solid',
            padding=5
        )
        self.template_label.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
=======
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
>>>>>>> new-branch-name
        
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
        
<<<<<<< HEAD
        # PDF Files Section
        pdf_frame = ttk.LabelFrame(self.root, text="PDF Files (Max 5) (Drag & Drop Enabled)", padding=10)
        pdf_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        pdf_frame.grid_columnconfigure(0, weight=1)
        
        self.pdf_label = ttk.Label(
            pdf_frame, 
            text="Drag & drop or click 'Select PDFs' to choose PDF files",
            background='white',
            relief='solid',
            padding=5
        )
        self.pdf_label.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
=======
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
>>>>>>> new-branch-name
        
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
        
<<<<<<< HEAD
        # Open Output Button
        self.open_output_btn = ttk.Button(
            button_frame,
            text="Open Output File",
            command=self._open_output_file,
            state="disabled"
        )
        self.open_output_btn.grid(row=0, column=1, padx=5)

    def _setup_drag_drop(self):
        """Setup drag and drop bindings"""
        # Get the window handle
        hwnd = self.root.winfo_id()
        
        # Set up the window procedure
        old_win_proc = win32gui.SetWindowLong(
            hwnd,
            win32con.GWL_WNDPROC,
            self._window_proc
        )
        
        # Store the old window procedure
        self.old_win_proc = old_win_proc
        
        # Visual feedback for drag and drop zones
        for widget in (self.template_label, self.pdf_label):
            widget.bind('<Enter>', lambda e, w=widget: self._on_drag_enter(w))
            widget.bind('<Leave>', lambda e, w=widget: self._on_drag_leave(w))

    def _get_dropped_files(self, hdrop) -> List[str]:
        """Get the list of dropped files using shell32"""
        result = []
        buf_size = 1024
        buf = create_unicode_buffer(buf_size)
        
        # Get count of files dropped
        count = windll.shell32.DragQueryFileW(c_uint(hdrop), -1, None, 0)
        
        # Get file paths
        for i in range(count):
            # Get required buffer size
            windll.shell32.DragQueryFileW(c_uint(hdrop), i, buf, buf_size)
            result.append(buf.value)
            
        return result

    def _window_proc(self, hwnd, msg, wparam, lparam):
        """Handle Windows messages"""
        try:
            if msg == win32con.WM_DROPFILES:
                try:
                    # Get drop point coordinates (in screen coordinates)
                    point = DWORD()
                    windll.shell32.DragQueryPoint(c_uint(wparam), point)
                    
                    # Convert screen coordinates to window coordinates
                    x = point.value & 0xFFFF  # Lower word
                    y = point.value >> 16     # Upper word
                    
                    files = self._get_dropped_files(wparam)
                    
                    # Convert window coordinates to widget coordinates
                    widget = self.root.winfo_containing(x, y)
                    
                    # Process the dropped files
                    if widget == self.template_label or widget in self.template_label.winfo_children():
                        for file in files:
                            if file.lower().endswith('.xlsx'):
                                self._set_template(file)
                                break
                            else:
                                messagebox.showwarning("Warning", "Please drop an Excel file (.xlsx)")
                    elif widget == self.pdf_label or widget in self.pdf_label.winfo_children():
                        for file in files:
                            if file.lower().endswith('.pdf'):
                                self._add_pdf(file)
                            else:
                                messagebox.showwarning("Warning", "Please drop only PDF files")
                finally:
                    # Always ensure DragFinish is called
                    try:
                        windll.shell32.DragFinish(wparam)
                    except:
                        pass
                return 0
                
            return win32gui.CallWindowProc(self.old_win_proc, hwnd, msg, wparam, lparam)
        except Exception as e:
            messagebox.showerror("Error", f"Error handling drop: {str(e)}")
            return win32gui.CallWindowProc(self.old_win_proc, hwnd, msg, wparam, lparam)

    def _on_drag_enter(self, widget):
        """Visual feedback when dragging over a drop zone"""
        widget.configure(relief='groove', background='#e6f3ff')

    def _on_drag_leave(self, widget):
        """Reset visual feedback when leaving drop zone"""
        widget.configure(relief='solid', background='white')
=======
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
>>>>>>> new-branch-name

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
<<<<<<< HEAD
        if len(self.pdf_files) >= 5:  # Updated limit to 5
            messagebox.showwarning("Warning", "Maximum 5 PDF files allowed")
=======
        print(f"Adding PDF file: {path}")  # Debug print
        if len(self.pdf_files) >= 5:
            QMessageBox.warning(self, "Maximum Files", "Maximum 5 PDF files allowed!")
>>>>>>> new-branch-name
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
<<<<<<< HEAD
            output_path = self.process_callback(self.template_file, self.pdf_files)
            if output_path:
                self.output_file = output_path
                messagebox.showinfo("Success", f"Financial analysis completed successfully!\nOutput saved to: {Path(output_path).name}")
                self.open_output_btn.configure(state="normal")
=======
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
>>>>>>> new-branch-name
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