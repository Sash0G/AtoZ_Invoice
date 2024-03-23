py -m PyInstaller --windowed --hidden-import=tkinter --hidden-import=tkinter.filedialog --hidden-import=tkinter.font --hidden-import=distutils --hidden-import=customtkinter --icon=Images/ship.ico ./AtoZ_Invoice.py
xcopy /s /i .\Images     .\dist\AtoZ_Invoice\_internal\Images
xcopy .\config.txt     .\dist\AtoZ_Invoice\_internal
xcopy .\data.db     .\dist\AtoZ_Invoice\_internal
xcopy .\template.xlsx     .\dist\AtoZ_Invoice\_internal
RMDIR /S /Q .\build
signer /sign FlyTechVideos .\dist\AtoZ_Invoice\AtoZ_Invoice.exe
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YYYY=%dt:~0,4%" & set "MM=%dt:~4,2%" & set "DD=%dt:~6,2%"
rename .\dist\AtoZ_Invoice AtoZ_Invoice_%DD%.%MM%.%YYYY%
move .\dist\AtoZ_Invoice_%DD%.%MM%.%YYYY% .\..\Releases
RMDIR /S /Q .\dist
@echo zdr 123
pause
