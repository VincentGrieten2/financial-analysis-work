import os
from JRR_automate import process_pdfs
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('parser.log'),
        logging.StreamHandler()
    ]
)

def test_csv_processing():
    try:
        # Get current directory
        current_dir = os.getcwd()
        
        # Define input files
        template_path = os.path.join(current_dir, "LRM_analyse_template_FD.xlsx")
        csv_path = os.path.join(current_dir, "2024-00505064.csv")
        
        # Verify files exist
        if not os.path.exists(template_path):
            logging.error(f"Template file not found at {template_path}")
            return
        if not os.path.exists(csv_path):
            logging.error(f"CSV file not found at {csv_path}")
            return
            
        logging.info(f"Processing files:")
        logging.info(f"Template: {template_path}")
        logging.info(f"CSV: {csv_path}")
        
        # Process the files
        try:
            output_path = process_pdfs(template_path, [csv_path])
            
            if output_path and os.path.exists(output_path):
                logging.info(f"Success! Output saved to: {output_path}")
            else:
                logging.error("Processing failed or output file not found")
        except Exception as e:
            logging.error(f"Error during file processing: {str(e)}")
            raise
            
    except Exception as e:
        logging.error(f"Error during test execution: {str(e)}")
        raise

if __name__ == "__main__":
    test_csv_processing() 