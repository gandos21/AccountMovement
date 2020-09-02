@echo off

set FILE_NAME=AcMove

echo Compiling Py script to EXE...
pyinstaller --onefile -w %FILE_NAME%.py



rem -- renaming from exe to dll no longer required as we have stopped using Excel macros

rem Delete existing .dll file and then rename .exe to .dll
rem del /f .\dist\%FILE_NAME%.dll
rem Wait a second to exe file creation to complete, before renaming it
rem timeout 1
rem rename .\dist\%FILE_NAME%.exe %FILE_NAME%.dll

pause
