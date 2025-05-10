# Financial Analysis Tool

A tool for extracting financial data from PDF statements and organizing it into standardized Excel format.

## Release Process

### To create a new release:

1. Update the version in `version.py`:
   ```python
   VERSION = "1.0.4"  # Update this number
   ```

2. Commit and push your changes:
   ```bash
   git add version.py
   git commit -m "Bump version to 1.0.4"
   git push origin main
   ```

3. Create and push a new tag:
   ```bash
   git tag v1.0.4
   git push origin v1.0.4
   ```

4. Build the release files locally:
   ```bash
   # Install requirements if not already installed
   pip install -r requirements.txt
   pip install pyinstaller

   # Create the executable
   pyinstaller --onefile --noconsole --name="Financial_Analysis" --add-data "version.py;." --add-data "gui.py;." run_financial_analysis.py

   # Create distribution folder and copy files
   mkdir dist\Financial_Analysis
   copy dist\Financial_Analysis.exe dist\Financial_Analysis\

   # Create ZIP file
   Compress-Archive -Path dist\Financial_Analysis\* -DestinationPath dist\Financial_Analysis_v1.0.4.zip -Force
   ```

5. Create the release on GitHub:
   - Go to the [Releases page](https://github.com/VincentGrieten2/financial-analysis-work/releases)
   - Click "Draft a new release"
   - Choose the tag you just created (e.g., "v1.0.4")
   - Set the release title (e.g., "Release v1.0.4")
   - Add release notes describing what's new
   - Upload the ZIP file from your `dist` folder
   - Click "Publish release"

## For Developers

### Setting Up Development Environment

1. Clone the repository
2. Create a virtual environment:
   ```
   python -m venv venv
   venv\Scripts\activate
   ```
3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

### Building Locally

Run `setup.bat` to build the executable locally. 