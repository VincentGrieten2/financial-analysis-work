import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import List, Callable, Tuple
import subprocess
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
        self.process_callback = process_callback
        self.pdf_files: List[str] = []
        self.template_file: str = ""
        self.output_file: str = ""
        
        # Create the main window
        self.root = tk.Tk()
        self.root.title("Financial Analysis Tool")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
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
        
        template_btn_frame = ttk.Frame(template_frame)
        template_btn_frame.grid(row=1, column=0, sticky="ew")
        template_btn_frame.grid_columnconfigure(1, weight=1)
        
        self.select_template_btn = ttk.Button(
            template_btn_frame, 
            text="Select Template", 
            command=self._select_template
        )
        self.select_template_btn.grid(row=0, column=0, padx=5)
        
        self.clear_template_btn = ttk.Button(
            template_btn_frame,
            text="Clear",
            command=self._clear_template,
            state="disabled"
        )
        self.clear_template_btn.grid(row=0, column=2, padx=5)
        
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
        
        pdf_btn_frame = ttk.Frame(pdf_frame)
        pdf_btn_frame.grid(row=1, column=0, sticky="ew")
        pdf_btn_frame.grid_columnconfigure(1, weight=1)
        
        self.select_pdf_btn = ttk.Button(
            pdf_btn_frame,
            text="Select PDFs",
            command=self._select_pdfs
        )
        self.select_pdf_btn.grid(row=0, column=0, padx=5)
        
        self.clear_pdfs_btn = ttk.Button(
            pdf_btn_frame,
            text="Clear All",
            command=self._clear_pdfs,
            state="disabled"
        )
        self.clear_pdfs_btn.grid(row=0, column=2, padx=5)
        
        # PDF List with Frame
        list_frame = ttk.Frame(self.root)
        list_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        
        # PDF List
        self.pdf_listbox = tk.Listbox(
            list_frame, 
            selectmode=tk.SINGLE, 
            height=6,
            background='white',
            relief='solid'
        )
        self.pdf_listbox.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbar for PDF list
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.pdf_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.pdf_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Process and Open Buttons Frame
        button_frame = ttk.Frame(self.root)
        button_frame.grid(row=3, column=0, pady=10)
        
        # Process Button
        self.process_btn = ttk.Button(
            button_frame,
            text="Process Files",
            command=self._process_files,
            state="disabled"
        )
        self.process_btn.grid(row=0, column=0, padx=5)
        
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

    def _select_template(self):
        """Open file dialog to select template"""
        file_path = filedialog.askopenfilename(
            title="Select Excel Template",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self._set_template(file_path)
    
    def _select_pdfs(self):
        """Open file dialog to select PDFs"""
        files = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF files", "*.pdf")]
        )
        for file in files:
            self._add_pdf(file)
    
    def _set_template(self, path: str):
        """Set the template file"""
        self.template_file = path
        self.template_label.configure(text=f"Template: {Path(path).name}")
        self.clear_template_btn.configure(state="normal")
        self._update_process_button()
    
    def _clear_template(self):
        """Clear the template selection"""
        self.template_file = ""
        self.template_label.configure(text="Click 'Select Template' to choose an Excel file")
        self.clear_template_btn.configure(state="disabled")
        self._update_process_button()
        # Reset output file when clearing template
        self.output_file = ""
        self.open_output_btn.configure(state="disabled")
    
    def _add_pdf(self, path: str):
        """Add a PDF file to the list"""
        if len(self.pdf_files) >= 5:  # Updated limit to 5
            messagebox.showwarning("Warning", "Maximum 5 PDF files allowed")
            return
        
        if path not in self.pdf_files:
            self.pdf_files.append(path)
            self.pdf_listbox.insert(tk.END, Path(path).name)
            self.clear_pdfs_btn.configure(state="normal")
            self._update_process_button()
    
    def _clear_pdfs(self):
        """Clear all PDF selections"""
        self.pdf_files.clear()
        self.pdf_listbox.delete(0, tk.END)
        self.clear_pdfs_btn.configure(state="disabled")
        self._update_process_button()
        # Reset output file when clearing PDFs
        self.output_file = ""
        self.open_output_btn.configure(state="disabled")
    
    def _update_process_button(self):
        """Update the state of the process button"""
        if self.template_file and self.pdf_files:
            self.process_btn.configure(state="normal")
        else:
            self.process_btn.configure(state="disabled")
    
    def _process_files(self):
        """Process the selected files"""
        try:
            output_path = self.process_callback(self.template_file, self.pdf_files)
            if output_path:
                self.output_file = output_path
                messagebox.showinfo("Success", f"Financial analysis completed successfully!\nOutput saved to: {Path(output_path).name}")
                self.open_output_btn.configure(state="normal")
            else:
                messagebox.showerror("Error", "An error occurred during processing. Check parser.log for details.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def _open_output_file(self):
        """Open the output Excel file"""
        if not self.output_file or not os.path.exists(self.output_file):
            messagebox.showerror("Error", "Output file not found!")
            return
        
        try:
            if os.name == 'nt':  # Windows
                os.startfile(self.output_file)
            else:  # macOS and Linux
                subprocess.run(['xdg-open' if os.name == 'posix' else 'open', self.output_file])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {str(e)}")
    
    def run(self):
        """Start the GUI"""
        self.root.mainloop() 