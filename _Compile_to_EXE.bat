@echo off

set FILE_NAME=AcMove

echo Compiling Py script to EXE...
pyinstaller --onefile %FILE_NAME%.py

rem Delete existing .dll file and then rename .exe to .dll
del /f .\dist\%FILE_NAME%.dll
rem Wait a second to exe file creation to complete, before renaming it
timeout 1
rename .\dist\%FILE_NAME%.exe %FILE_NAME%.dll

pause
