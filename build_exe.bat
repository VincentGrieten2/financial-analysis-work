@echo off
pyinstaller --onefile --noconsole --name="Financial_Analysis" ^
  --add-data "version.py;." ^
  --add-data "gui.py;." ^
  run_financial_analysis.py

mkdir dist\Financial_Analysis
copy dist\Financial_Analysis.exe dist\Financial_Analysis\ 