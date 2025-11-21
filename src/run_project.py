"""
PDF-to-Spreadsheet Converter - Simple Project Runner
"""

import os
import sys
import subprocess
from pathlib import Path
import importlib
import importlib.util

# Determine project root (one level up from src/)
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Required packages (specs)
PACKAGES = [
    "pandas>=1.5.0",
    "pdfplumber>=0.9.0",
    "openpyxl>=3.1.0"
]

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

def _base_name(spec: str) -> str:
    return spec.split('>=')[0].strip()

def get_missing_packages(packages):
    """Return list of package specs that are not importable in this env."""
    missing = []
    for spec in packages:
        name = _base_name(spec)
        try:
            importlib.import_module(name)
        except Exception:
            missing.append(spec)
    return missing

def setup_environment(packages_to_install):
    """Install only the packages in packages_to_install (list of specs).
    If the list is empty, nothing is installed.
    """
    if not packages_to_install:
        print("All required packages are already installed.")
        return True
    print("Installing required packages...")
    failed_packages = []
    for spec in packages_to_install:
        pkg_name = _base_name(spec)
        print(f"Installing {pkg_name}...")
        if not install_package(spec):
            failed_packages.append(pkg_name)
    if failed_packages:
        print(f"Failed to install: {', '.join(failed_packages)}")
        print("Please install manually using: pip install <package_name>")
        return False
    print("Packages installed successfully")
    return True

def create_directories():
    """Create necessary project directories"""
    directories = [
        PROJECT_ROOT / "data" / "input",
        PROJECT_ROOT / "data" / "output"
    ]
    for directory in directories:
        Path(directory).mkdir(parents=True, exist_ok=True)

def run_converter():
    """Run the PDF converter: list PDFs in data/input, prompt for selection, and run converter."""
    input_dir = PROJECT_ROOT / "data" / "input"
    if not input_dir.exists():
        print(f"Error: input directory '{input_dir}' not found")
        return False
    pdf_files = [p for p in sorted(input_dir.iterdir()) if p.is_file() and p.suffix.lower() == ".pdf"]
    if not pdf_files:
        print(f"No PDF files found in '{input_dir}'")
        return True

    print("\n🚀 PDF-to-Spreadsheet Converter")
    print("Available PDF files:")
    for idx, p in enumerate(pdf_files, 1):
        print(f"  [{idx}] {p.name}")

    while True:
        choice = input(f"Select a PDF to convert [1-{len(pdf_files)}]: ").strip()
        if not choice:
            print("❌ No input provided. Please enter a number.")
            continue
        try:
            idx = int(choice)
            if not (1 <= idx <= len(pdf_files)):
                print("❌ Invalid selection. Try again.")
                continue
        except ValueError:
            print("❌ Invalid input. Please enter a number.")
            continue
        selected = pdf_files[idx - 1]
        break

    converter_path = PROJECT_ROOT / "src" / "pdf_converter.py"
    if not converter_path.exists():
        print(f"Error: converter script not found at {converter_path}")
        return False

    print(f"🔍 Processing PDF: {selected.name}")
    # Attempt to load and call the converter module directly
    try:
        spec = importlib.util.spec_from_file_location("pdf_converter_module", str(converter_path))
        if spec is None or spec.loader is None:
            raise RuntimeError("Failed to load converter module spec")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)

        # Prefer calling main with the selected file if available
        if hasattr(module, "main"):
            try:
                module.main(str(selected))
                return True
            except TypeError:
                # main() may take no args; call without
                module.main()
                return True

        # Common alternative function names
        for func_name in ("convert_pdf", "convert_file", "process_pdf", "convert"):
            if hasattr(module, func_name):
                getattr(module, func_name)(str(selected))
                return True

        # If there is a PDFConverter class, try to use it
        if hasattr(module, "PDFConverter"):
            try:
                conv = module.PDFConverter()
                for meth in ("convert", "process", "run", "convert_file"):
                    if hasattr(conv, meth):
                        getattr(conv, meth)(str(selected))
                        return True
            except Exception:
                pass

        # Fallback: run as subprocess
        subprocess.run([sys.executable, str(converter_path), str(selected)], check=True, cwd=str(PROJECT_ROOT))
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error running converter as subprocess: {e}")
        return False
    except Exception as e:
        print(f"Error invoking converter module: {e}")
        # final fallback to subprocess
        try:
            subprocess.run([sys.executable, str(converter_path), str(selected)], check=True, cwd=str(PROJECT_ROOT))
            return True
        except Exception as e2:
            print(f"Final fallback failed: {e2}")
            return False

def main():
    print("PDF-to-Spreadsheet Converter Setup")
    print("-" * 40)
    
    if not check_python_version():
        sys.exit(1)
    
    create_directories()
    
    # Install only missing packages
    missing = get_missing_packages(PACKAGES)
    if not setup_environment(missing):
        print("Setup failed. Please install dependencies manually.")
        sys.exit(1)
    
    print("\nRunning PDF converter...")
    if not run_converter():
        print("Converter failed to run")
        return
    
    print("\nSetup and conversion completed")

if __name__ == "__main__":
    main()