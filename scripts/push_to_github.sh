#!/usr/bin/env sh
set -e

echo "Enter remote Git repository URL (e.g. https://github.com/everBlack0/Automated-PDF-to-spreadsheet-Converter.git):"
read REMOTE_URL
if [ -z "$REMOTE_URL" ]; then
  echo "No remote URL provided. Exiting."
  exit 1
fi

if [ ! -d .git ]; then
  git init
fi

echo "Adding files..."
git add .

echo "Committing..."
git commit -m "Initial commit: PDF-to-Spreadsheet Converter" || echo "Nothing to commit"

if git remote | grep origin > /dev/null 2>&1; then
  git remote remove origin
fi

echo "Adding remote..."
git remote add origin "$REMOTE_URL"

echo "Pushing to remote (main)..."
git branch -M main

git push -u origin main

echo "Done."
