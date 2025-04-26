# Financial Analysis Tool

A tool for extracting financial data from PDF statements and organizing it into standardized Excel format.

## Continuous Integration/Continuous Deployment (CI/CD)

This project uses GitHub Actions for automated building and releasing. Here's how to use it:

### To create a new release:

1. Update the version in `version.py`:
   ```python
   VERSION = "1.0.1"  # Update this number
   ```

2. Commit and push your changes:
   ```bash
   git add version.py
   git commit -m "Bump version to 1.0.1"
   git push origin main
   ```

3. Create and push a new tag:
   ```bash
   git tag v1.0.1
   git push origin v1.0.1
   ```

4. The GitHub Actions workflow will automatically:
   - Build the executable
   - Create a draft release
   - Upload the distribution ZIP file

5. Go to GitHub Releases, review the draft release, add release notes, and publish it.

### CI/CD Workflow Details

The workflow defined in `.github/workflows/build.yml` executes when a new tag starting with 'v' is pushed. It:

1. Sets up a Windows environment with Python
2. Installs all dependencies
3. Builds the executable using PyInstaller
4. Creates a release ZIP with the application and necessary files
5. Creates a draft GitHub release
6. Uploads the ZIP file as an asset

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