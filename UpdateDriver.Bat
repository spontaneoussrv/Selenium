@echo off
setlocal

REM Define paths
set "currentFolder=%~dp0"
set "seleniumFolder=C:\Program files\seleniumBasic"

REM Copy msedgedriver.exe to selenium folder
copy "%currentFolder%msedgedriver.exe" "%seleniumFolder%"

REM Check if edgedriver.exe exists and delete it
if exist "%seleniumFolder%\edgedriver.exe" (
    del "%seleniumFolder%\edgedriver.exe"
)

REM Rename msedgedriver.exe to edgedriver.exe
rename "%seleniumFolder%\msedgedriver.exe" "edgedriver.exe"

echo Operation completed successfully.
endlocal
pause
