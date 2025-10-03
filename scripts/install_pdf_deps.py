#!/usr/bin/env python3
"""
install_pdf_deps.py

Automated script to install PDF dependencies with proper error handling
and verification.

Usage:
  python scripts/install_pdf_deps.py [--user] [--verbose] [--force]
"""

import sys
import subprocess
import argparse
from pathlib import Path

def run_command(cmd, description, check=True, capture_output=True):
    """Run a command with proper error handling."""
    print(f"🔧 {description}...")
    
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            check=check,
            capture_output=capture_output,
            text=True
        )
        
        if capture_output and result.stdout:
            print(result.stdout)
            
        return True, result
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to {description.lower()}")
        if capture_output and e.stderr:
            print(f"Error: {e.stderr}")
        return False, e

def check_python_version():
    """Check if Python version is compatible."""
    if sys.version_info < (3, 8):
        print("❌ Python 3.8+ is required")
        return False
    print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor} detected")
    return True

def install_requirements(user=False, force=False):
    """Install from requirements file."""
    req_file = Path("requirements_pdf.txt")
    
    if not req_file.exists():
        print("❌ requirements_pdf.txt not found")
        return False
    
    cmd = "pip install"
    if user:
        cmd += " --user"
    if force:
        cmd += " --force-reinstall"
    cmd += f" -r {req_file}"
    
    success, _ = run_command(cmd, "Installing from requirements_pdf.txt")
    return success

def install_individual_packages(user=False, force=False):
    """Install packages individually."""
    packages = [
        "PyMuPDF>=1.22.0",
        "pdfplumber>=0.9.0", 
        "pdfminer.six>=20221105",
        "pypdf>=3.0.0",
        "Pillow>=9.0.0"
    ]
    
    cmd = "pip install"
    if user:
        cmd += " --user"
    if force:
        cmd += " --force-reinstall"
    cmd += " " + " ".join(packages)
    
    success, _ = run_command(cmd, "Installing individual packages")
    return success

def main():
    """Main installation function."""
    parser = argparse.ArgumentParser(description="Install PDF dependencies")
    parser.add_argument('--user', '-u', action='store_true', help='Install to user directory')
    parser.add_argument('--verbose', '-v', action='store_true', help='Verbose output')
    parser.add_argument('--force', '-f', action='store_true', help='Force reinstall')
    parser.add_argument('--individual', '-i', action='store_true', help='Install packages individually')
    
    args = parser.parse_args()
    
    print("🚀 PDF Dependencies Installation Script")
    print("=" * 50)
    
    # Check Python version
    if not check_python_version():
        return 1
    
    # Check if we're in a virtual environment
    if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
        print("✅ Virtual environment detected")
    else:
        print("⚠️  No virtual environment detected - consider using --user flag")
    
    # Install dependencies
    print(f"\n📦 Installing PDF dependencies...")
    print("-" * 30)
    
    if args.individual:
        success = install_individual_packages(args.user, args.force)
    else:
        success = install_requirements(args.user, args.force)
        
        if not success:
            print("\n🔄 Trying individual package installation...")
            success = install_individual_packages(args.user, args.force)
    
    if not success:
        print("\n❌ Installation failed")
        print("\n💡 Alternative installation methods:")
        print("   1. Use virtual environment: python -m venv pdf-env")
        print("   2. Install individually: pip install --user PyMuPDF pdfplumber")
        print("   3. Check system Python permissions")
        return 1
    
    print("\n✅ Installation completed!")
    
    # Run verification
    print(f"\n🔍 Verifying installation...")
    print("-" * 30)
    
    verify_cmd = "python scripts/verify_pdf_deps.py"
    if args.verbose:
        verify_cmd += " --verbose"
    
    success, result = run_command(verify_cmd, "Verifying dependencies", check=False)
    
    if success and result.returncode == 0:
        print("\n🎉 All dependencies are properly installed and working!")
        return 0
    else:
        print("\n⚠️  Installation succeeded but verification failed")
        print("   Check the verification output above for details")
        return 1

if __name__ == "__main__":
    sys.exit(main())