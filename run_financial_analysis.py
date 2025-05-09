import os
import sys
from pathlib import Path
from JRR_automate import process_pdfs
from updater import get_update_status
from gui import FinancialAnalysisGUI, QApplication
import logging

# Configure logging
handlers = [logging.FileHandler('parser.log')]
# Only add StreamHandler if not running as executable
if not getattr(sys, 'frozen', False):
    handlers.append(logging.StreamHandler())

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=handlers
)

def process_files(template_path: str, pdf_paths: list) -> bool:
    """
    Process the selected files
    :param template_path: Path to the Excel template
    :param pdf_paths: List of paths to PDF files
    :return: True if processing was successful, False otherwise
    """
    try:
        # Log the processing attempt
        logging.info(f"Processing files with template: {template_path}")
        logging.info(f"PDF files: {pdf_paths}")
        
        # Process the files
        return process_pdfs(template_path, pdf_paths)
        
    except Exception as e:
        logging.error(f"Error processing files: {str(e)}")
        return False

def main():
    try:
        # Check for updates
        update_status = get_update_status()
        logging.info(update_status)
        
        # Initialize QApplication
        app = QApplication(sys.argv)
        
        # Create and run the GUI
        gui = FinancialAnalysisGUI(process_files)
        gui.show()  # Show the window
        
        # Start the event loop
        sys.exit(app.exec())
        
    except Exception as e:
        logging.error(f"Application error: {str(e)}")
        # Only print to console if not running as executable
        if not getattr(sys, 'frozen', False):
            print(f"\nAn error occurred: {str(e)}")
            print("Check parser.log for more details")
            input("\nPress Enter to exit...")

if __name__ == "__main__":
    main() 