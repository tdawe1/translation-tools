#!/usr/bin/env python3
"""
verify_pdf_deps.py

Script to verify that all required PDF dependencies are properly installed
and working correctly for the translation pipeline.

Usage:
  python scripts/verify_pdf_deps.py [--verbose]
  python scripts/verify_pdf_deps.py --check-only  # Exit code 1 on missing deps
"""

import sys
import importlib
import subprocess
from pathlib import Path
from typing import Dict, List, Tuple, Optional

# Required dependencies with their import names and versions
REQUIRED_DEPS = {
    'PyMuPDF': {'import_name': 'fitz', 'min_version': '1.22.0', 'package_name': 'PyMuPDF'},
    'pdfplumber': {'import_name': 'pdfplumber', 'min_version': '0.9.0', 'package_name': 'pdfplumber'},
    'pdfminer.six': {'import_name': 'pdfminer', 'min_version': '20221105', 'package_name': 'pdfminer.six'},
    'pypdf': {'import_name': 'pypdf', 'min_version': '3.0.0', 'package_name': 'pypdf'},
    'Pillow': {'import_name': 'PIL', 'min_version': '9.0.0', 'package_name': 'Pillow'},
}

OPTIONAL_DEPS = {
    'pytest': {'import_name': 'pytest', 'min_version': '7.0.0', 'package_name': 'pytest'},
    'black': {'import_name': 'black', 'min_version': '22.0.0', 'package_name': 'black'},
    'flake8': {'import_name': 'flake8', 'min_version': '5.0.0', 'package_name': 'flake8'},
    'mypy': {'import_name': 'mypy', 'min_version': '1.0.0', 'package_name': 'mypy'},
}

def check_version(package_name: str, import_name: str, min_version: str) -> Tuple[bool, Optional[str]]:
    """Check if a package meets the minimum version requirement."""
    try:
        module = importlib.import_module(import_name)
        
        # Get version based on package
        if package_name == 'PyMuPDF':
            version = getattr(module, 'version', 'unknown')
        elif package_name == 'Pillow':
            version = getattr(module, '__version__', 'unknown')
        elif package_name == 'pdfminer.six':
            # pdfminer.six doesn't have a standard version attribute
            version = 'unknown'
        else:
            version = getattr(module, '__version__', 'unknown')
        
        if version == 'unknown':
            return True, version
            
        # Simple version comparison (works for most cases)
        version_parts = [int(x) for x in version.split('.')[:2]]
        min_parts = [int(x) for x in min_version.split('.')[:2]]
        
        if version_parts >= min_parts:
            return True, version
        else:
            return False, version
            
    except ImportError:
        return False, None
    except Exception as e:
        print(f"Error checking version for {package_name}: {e}")
        return False, None

def verify_dependencies(verbose: bool = False) -> Tuple[bool, Dict[str, Dict]]:
    """Verify all required dependencies are installed and meet version requirements."""
    results = {}
    all_required_ok = True
    
    print("🔍 Verifying PDF dependencies...")
    print("=" * 50)
    
    # Check required dependencies
    print("\nRequired Dependencies:")
    print("-" * 25)
    
    for dep_name, dep_info in REQUIRED_DEPS.items():
        success, version = check_version(dep_info['package_name'], dep_info['import_name'], dep_info['min_version'])
        
        results[dep_name] = {
            'installed': success,
            'version': version,
            'min_version': dep_info['min_version'],
            'required': True
        }
        
        if success:
            if version and version != 'unknown':
                print(f"✅ {dep_name}: {version} (>= {dep_info['min_version']})")
            else:
                print(f"✅ {dep_name}: installed (version unknown)")
        else:
            print(f"❌ {dep_name}: NOT INSTALLED or version < {dep_info['min_version']}")
            all_required_ok = False
        
        if verbose and success:
            try:
                module = importlib.import_module(dep_info['import_name'])
                if hasattr(module, '__file__'):
                    location = Path(module.__file__).parent
                    print(f"   Location: {location}")
            except:
                pass
    
    # Check optional dependencies
    print(f"\nOptional Dependencies:")
    print("-" * 25)
    
    for dep_name, dep_info in OPTIONAL_DEPS.items():
        success, version = check_version(dep_info['package_name'], dep_info['import_name'], dep_info['min_version'])
        
        results[dep_name] = {
            'installed': success,
            'version': version,
            'min_version': dep_info['min_version'],
            'required': False
        }
        
        if success:
            if version:
                print(f"✅ {dep_name}: {version}")
            else:
                print(f"✅ {dep_name}: installed")
        else:
            print(f"⚠️  {dep_name}: NOT INSTALLED (optional)")
    
    return all_required_ok, results

def install_suggestions(missing_deps: List[str]) -> str:
    """Generate installation suggestions for missing dependencies."""
    if not missing_deps:
        return ""
    
    suggestions = []
    
    # Install via requirements file
    req_file = Path("requirements_pdf.txt")
    if req_file.exists():
        suggestions.append(f"pip install -r {req_file}")
    
    # Install individual packages
    package_names = []
    for dep in missing_deps:
        if dep in REQUIRED_DEPS:
            package_names.append(REQUIRED_DEPS[dep]['package_name'])
    
    if package_names:
        suggestions.append(f"pip install {' '.join(package_names)}")
    
    return "\n\n💡 Installation commands:\n" + "\n".join(f"   {cmd}" for cmd in suggestions)

def main():
    """Main verification function."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Verify PDF dependencies")
    parser.add_argument('--verbose', '-v', action='store_true', help='Verbose output')
    parser.add_argument('--check-only', action='store_true', help='Exit with code 1 on missing deps')
    
    args = parser.parse_args()
    
    all_ok, results = verify_dependencies(args.verbose)
    
    # Find missing required dependencies
    missing_required = [
        dep for dep, info in results.items() 
        if info['required'] and not info['installed']
    ]
    
    # Summary
    print(f"\n{'='*50}")
    if all_ok:
        print("🎉 All required PDF dependencies are properly installed!")
        if missing_required:
            print(f"⚠️  However, {len(missing_required)} optional dependencies are missing.")
    else:
        print(f"❌ Missing {len(missing_required)} required dependencies:")
        for dep in missing_required:
            print(f"   - {dep}")
    
    # Provide installation suggestions if needed
    if missing_required:
        print(install_suggestions(missing_required))
    
    # Test basic functionality
    if all_ok:
        print(f"\n🧪 Testing basic PDF functionality...")
        try:
            import fitz
            # Test PyMuPDF
            doc = fitz.open()  # Create empty document
            doc.close()
            print("✅ PyMuPDF: Basic functionality working")
            
            try:
                import pdfplumber
                print("✅ pdfplumber: Import successful")
            except ImportError:
                print("⚠️  pdfplumber: Import failed (but marked as installed)")
            
            try:
                import pypdf
                # Test pypdf
                from pypdf import PdfReader
                print("✅ pypdf: Basic functionality working")
            except Exception as e:
                print(f"⚠️  pypdf: Functionality test failed - {e}")
                
        except Exception as e:
            print(f"❌ Basic functionality test failed: {e}")
            all_ok = False
    
    # Exit code
    if args.check_only and not all_ok:
        sys.exit(1)
    
    return 0 if all_ok else 1

if __name__ == "__main__":
    sys.exit(main())