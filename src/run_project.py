"""
PDF-to-Spreadsheet Converter - Simple Project Runner
"""

import os
import sys
import subprocess
from pathlib import Path

def check_python_version():
    """Check if Python version is compatible"""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print(f"Python 3.8+ required. Current version: {version.major}.{version.minor}.{version.micro}")
        return False
    print(f"Python version: {version.major}.{version.minor}.{version.micro}")
    return True

def install_package(package):
    """Install a single package silently"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package], 
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except subprocess.CalledProcessError:
        return False

def setup_environment():
    """Setup project environment and install dependencies"""
    print("Installing required packages...")
    
    packages = [
        "pandas>=1.5.0",
        "pdfplumber>=0.9.0", 
        "openpyxl>=3.1.0"
    ]
    
    failed_packages = []
    
    for package in packages:
        package_name = package.split('>=')[0]
        print(f"Installing {package_name}...")
        
        if not install_package(package):
            failed_packages.append(package_name)
    
    if failed_packages:
        print(f"Failed to install: {', '.join(failed_packages)}")
        print("Please install manually using: pip install <package_name>")
        return False
    
    print("Packages installed successfully")
    return True

def create_directories():
    """Create necessary project directories"""
    directories = [
        os.path.join("data", "input"),
        os.path.join("data", "output")
    ]
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)

def run_converter():
    """Run the PDF converter"""
    converter_path = os.path.join("src", "pdf_converter.py")
    if not os.path.exists(converter_path):
        print(f"Error: {converter_path} not found")
        return False
    
    try:
        subprocess.run([sys.executable, converter_path], check=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error running converter: {e}")
        return False
    except KeyboardInterrupt:
        print("Operation cancelled")
        return False

def main():
    print("PDF-to-Spreadsheet Converter Setup")
    print("-" * 40)
    
    if not check_python_version():
        sys.exit(1)
    
    create_directories()
    
    if not setup_environment():
        print("Setup failed. Please install dependencies manually.")
        sys.exit(1)
    
    print("\nRunning PDF converter...")
    if not run_converter():
        print("Converter failed to run")
        return
    
    print("\nSetup and conversion completed")

if __name__ == "__main__":
    main()