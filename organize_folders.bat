@echo off
setlocal enabledelayedexpansion

:: Step 1: Get Today's Date in DD-MM-YYYY Format
for /f "tokens=2 delims==." %%a in ('wmic os get localdatetime /value') do set "datetime=%%a"
set "year=%datetime:~0,4%"
set "month=%datetime:~4,2%"
set "day=%datetime:~6,2%"
set "today=%day%-%month%-%year%"

:: Step 2: Create Today's Folder
mkdir "%today%" 2>nul

:: Step 3: Create Done, Exist, Incorrect
mkdir "%today%\Done" 2>nul
mkdir "%today%\Exist" 2>nul
mkdir "%today%\Incorrect" 2>nul

:: Step 4: Move ALL Subfolders to Done (Dynamic) and Rename them
for /d %%i in (*) do (
  if /i not "%%i"=="%today%" (
    set "oldname=%%i"
    set "newname=Done %%i"
    move "%%i" "%today%\Done\!newname!" >nul
  )
)

:: Step 5: Create Exist and Incorrect Folders with Correct Names
cd "%today%"
for /d %%j in (Done\*) do (
  set "folder=%%~nxj"
  set "originalName=!folder:Done =!"  
  mkdir "Exist\Exist !originalName!" 2>nul
  mkdir "Incorrect\Incorrect !originalName!" 2>nul
)

echo Folders Organized in "%today%"
pause