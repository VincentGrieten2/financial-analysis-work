@echo off
echo Reading version information...
for /f "tokens=2 delims='='" %%i in ('type version.py ^| findstr VERSION') do set VERSION=%%i
set VERSION=%VERSION:"=%
echo Current version: %VERSION%

echo Installing required packages...
pip install -r requirements.txt
pip install pyinstaller

echo Creating executable...
pyinstaller --onefile --noconsole --name="Financial_Analysis" ^
    --add-data "version.py;." ^
    --add-data "gui.py;." ^
    run_financial_analysis.py

echo Creating distribution package...
mkdir dist\Financial_Analysis
copy dist\Financial_Analysis.exe dist\Financial_Analysis\

echo Creating distribution ZIP...
powershell Compress-Archive -Path dist\Financial_Analysis\* -DestinationPath dist\Financial_Analysis_v%VERSION%.zip -Force

echo Cleaning up...
rmdir /s /q build
del Financial_Analysis.spec

echo Done! Distribution package created at dist\Financial_Analysis_v%VERSION%.zip
pause 