import os
import sys
from pathlib import Path
from JRR_automate import process_pdfs
from updater import get_update_status
from gui import FinancialAnalysisGUI
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

def process_files(template_path: str, pdf_paths: list) -> str:
    """
    Process the selected files
    :param template_path: Path to the Excel template
    :param pdf_paths: List of paths to PDF files
    :return: Path to the output file if successful, None otherwise
    """
    try:
        # Log the processing attempt
        logging.info(f"Processing files with template: {template_path}")
        logging.info(f"PDF files: {pdf_paths}")
        
        # Process the files
        success = process_pdfs(template_path, pdf_paths)
        if success:
            # Get the most recent output file
            files = [f for f in os.listdir() if f.startswith('financial_analysis_') and f.endswith('.xlsx')]
            if files:
                return max(files, key=lambda x: os.path.getctime(x))
        return None
        
    except Exception as e:
        logging.error(f"Error processing files: {str(e)}")
        return None

def main():
    try:
        # Check for updates
        update_status = get_update_status()
        logging.info(update_status)
        
        # Create and run the GUI
        gui = FinancialAnalysisGUI(process_files)
        gui.run()
        
    except Exception as e:
        logging.error(f"Application error: {str(e)}")
        # Only print to console if not running as executable
        if not getattr(sys, 'frozen', False):
            print(f"\nAn error occurred: {str(e)}")
            print("Check parser.log for more details")
            input("\nPress Enter to exit...")

if __name__ == "__main__":
    main() 