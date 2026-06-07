
cls
REM
python -m PyInstaller --add-data "browser;browser" --onedir --clean main.py

mkdir .\dist\downloadAxon\browser
xcopy .\browser .\dist\downloadAxon\browser /E /I /H /K /Y

cd .\dist
