@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

REM Push project to GitHub helper (Windows)
SET /P REMOTE_URL=Enter remote Git repository URL (e.g. https://github.com/everBlack0/Automated-PDF-to-spreadsheet-Converter.git): 

if "%REMOTE_URL%"=="" (
  echo No remote URL provided. Exiting.
  exit /b 1
)

echo Initializing git repository...
if not exist .git (
  git init
) else (
  echo .git already exists
)

echo Adding files...
git add .

echo Committing...
git commit -m "Initial commit: PDF-to-Spreadsheet Converter" || echo "Nothing to commit"

echo Adding remote...
git remote remove origin 2>nul
ngit remote add origin %REMOTE_URL%

echo Pushing to remote (main)...
rem create main branch if not exists
git branch -M main

git push -u origin main

echo Done.
