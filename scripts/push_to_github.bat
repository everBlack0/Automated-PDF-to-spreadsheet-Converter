@echo off
SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION

REM Push project to GitHub helper (Windows)
REM Usage: push_to_github.bat [remote_url]

if "%~1"=="" (
  set /p "REMOTE_URL=Enter remote Git repository URL (e.g. https://github.com/everBlack0/Automated-PDF-to-spreadsheet-Converter.git): "
) else (
  set "REMOTE_URL=%~1"
)

if "%REMOTE_URL%"=="" (
  echo No remote URL provided. Exiting.
  exit /b 1
)

REM Verify git is available
where git >nul 2>&1
if errorlevel 1 (
  echo Git not found in PATH. Please install Git and ensure it's available in the PATH.
  exit /b 1
)

echo Initializing git repository...
if not exist ".git" (
  git init
) else (
  echo .git already exists
)

echo Adding files...
git add .

echo Committing...
git commit -m "Initial commit: PDF-to-Spreadsheet Converter" 2>nul || echo Nothing to commit

echo Adding remote (origin = %REMOTE_URL%)...
rem remove existing origin if present
git remote remove origin 2>nul || rem ignore

git remote add origin "%REMOTE_URL%"

echo Pushing to remote (main)...
REM create/rename branch to main if needed
git branch -M main 2>nul || rem ignore

git push -u origin main

echo Done.
