#!/usr/bin/env python3
"""
Comprehensive test runner for PDF translation system.

Provides a unified interface for running different test suites and generating reports.
"""

import sys
import os
import json
import time
import argparse
from pathlib import Path
from datetime import datetime
import subprocess
import unittest

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

class TestRunner:
    """Unified test runner for PDF translation system."""
    
    def __init__(self):
        self.results = {}
        self.start_time = time.time()
        self.test_suites = {
            'quality': {
                'name': 'Quality Metrics Tests',
                'description': 'Tests quality thresholds and metrics enforcement',
                'command': ['python', '-m', 'pytest', str(project_root / 'tests/test_pdf_quality_metrics.py'), '-v', '--tb=short'],
                'required': True
            },
            'integration-quality': {
                'name': 'Integration Quality Tests', 
                'description': 'Tests quality metrics in integration context',
                'command': ['python', '-m', 'pytest', str(project_root / 'tests/test_pdf_integration.py::TestQualityMetricsEnforcement'), '-v', '--tb=short'],
                'required': True
            },
            'integration-errors': {
                'name': 'Error Handling Tests',
                'description': 'Tests error scenarios and edge cases',
                'command': ['python', '-m', 'pytest', str(project_root / 'tests/test_pdf_integration.py::TestPDFTranslationErrorScenarios'), '-v', '--tb=short'],
                'required': False
            },
            'integration-output': {
                'name': 'Output Generation Tests',
                'description': 'Tests CSV and audit report generation',
                'command': ['python', '-m', 'pytest', str(project_root / 'tests/test_pdf_integration.py::TestPDFTranslationOutputGeneration'), '-v', '--tb=short'],
                'required': False
            },
            'integration-e2e': {
                'name': 'End-to-End Tests',
                'description': 'Tests complete translation pipeline',
                'command': ['python', '-m', 'pytest', str(project_root / 'tests/test_pdf_integration.py::TestPDFEndToEndTranslation'), '-v', '--tb=short'],
                'required': False
            },
            'validation': {
                'name': 'Test Data Validation',
                'description': 'Validates sample data and test fixtures',
                'command': ['python', '-m', 'pytest', str(project_root / 'tests/test_pdf_quality_metrics.py::TestPDFIntegrationValidation'), '-v', '--tb=short'],
                'required': True
            }
        }
    
    def run_test_suite(self, suite_name):
        """Run a specific test suite."""
        if suite_name not in self.test_suites:
            print(f"❌ Unknown test suite: {suite_name}")
            return False
        
        suite = self.test_suites[suite_name]
        print(f"\n🔍 Running {suite['name']}")
        print(f"   Description: {suite['description']}")
        print(f"   Required: {'Yes' if suite['required'] else 'No'}")
        print("   " + "="*60)
        
        start_time = time.time()
        
        try:
            # Run the test command
            result = subprocess.run(
                suite['command'],
                cwd=project_root,
                capture_output=True,
                text=True,
                timeout=300  # 5 minute timeout
            )
            
            end_time = time.time()
            duration = end_time - start_time
            
            # Store results
            self.results[suite_name] = {
                'success': result.returncode == 0,
                'duration': duration,
                'output': result.stdout,
                'error': result.stderr,
                'required': suite['required']
            }
            
            # Print results
            if result.returncode == 0:
                print(f"✅ {suite['name']} - PASSED ({duration:.2f}s)")
            else:
                print(f"❌ {suite['name']} - FAILED ({duration:.2f}s)")
                if result.stderr:
                    print(f"   Error: {result.stderr.strip()}")
            
            return result.returncode == 0
            
        except subprocess.TimeoutExpired:
            print(f"⏰ {suite['name']} - TIMEOUT (exceeded 5 minutes)")
            self.results[suite_name] = {
                'success': False,
                'duration': 300,
                'output': '',
                'error': 'Test timeout',
                'required': suite['required']
            }
            return False
            
        except Exception as e:
            print(f"💥 {suite['name']} - ERROR: {str(e)}")
            self.results[suite_name] = {
                'success': False,
                'duration': 0,
                'output': '',
                'error': str(e),
                'required': suite['required']
            }
            return False
    
    def run_all_tests(self):
        """Run all test suites."""
        print("🚀 Running PDF Translation Integration Tests")
        print("=" * 60)
        
        # Generate test data first
        print("\n📄 Generating test data...")
        try:
            subprocess.run([
                'python', 'tests/create_sample_pdfs.py'
            ], cwd=project_root, check=True, capture_output=True)
            print("✅ Test data generated successfully")
        except subprocess.CalledProcessError as e:
            print(f"⚠️  Test data generation failed: {e}")
        
        # Run all test suites
        total_suites = len(self.test_suites)
        passed_suites = 0
        
        for suite_name in self.test_suites:
            if self.run_test_suite(suite_name):
                passed_suites += 1
        
        return passed_suites, total_suites
    
    def generate_report(self):
        """Generate comprehensive test report."""
        total_duration = time.time() - self.start_time
        
        print("\n" + "=" * 60)
        print("📊 TEST RESULTS SUMMARY")
        print("=" * 60)
        
        # Calculate overall stats
        total_tests = len(self.results)
        passed_tests = sum(1 for r in self.results.values() if r['success'])
        required_tests = sum(1 for r in self.results.values() if r['required'])
        passed_required = sum(1 for r in self.results.values() if r['success'] and r['required'])
        
        # Overall status
        overall_success = passed_required == required_tests
        
        print(f"Overall Status: {'✅ PASS' if overall_success else '❌ FAIL'}")
        print(f"Total Duration: {total_duration:.2f}s")
        print(f"Test Suites: {passed_tests}/{total_tests} passed")
        print(f"Required Tests: {passed_required}/{required_tests} passed")
        print()
        
        # Individual suite results
        print("Individual Test Suite Results:")
        print("-" * 40)
        
        for suite_name, result in self.results.items():
            suite = self.test_suites[suite_name]
            status = "✅ PASS" if result['success'] else "❌ FAIL"
            required = " (Required)" if suite['required'] else ""
            print(f"{suite['name']}: {status}{required} ({result['duration']:.2f}s)")
        
        # Quality metrics summary
        print("\n" + "Quality Metrics Summary:")
        print("-" * 40)
        
        quality_metrics = {
            'Residual Japanese': '≤2%',
            'Layout Integrity': '≥95%', 
            'Cache Efficiency': '≥80%',
            'Translation Completeness': '≥85%',
            'Overall Quality': '≥80%'
        }
        
        for metric, threshold in quality_metrics.items():
            print(f"• {metric}: {threshold}")
        
        # Recommendations
        if not overall_success:
            print("\n" + "🔧 Recommendations:")
            failed_required = [name for name, r in self.results.items() if not r['success'] and r['required']]
            if failed_required:
                print("• Fix required test failures:")
                for name in failed_required:
                    print(f"  - {self.test_suites[name]['name']}")
            
            print("• Review error messages above for specific failure details")
            print("• Check dependencies and test environment setup")
            print("• Ensure sample data is properly generated")
        
        # Generate JSON report
        report_data = {
            'timestamp': datetime.now().isoformat(),
            'overall_success': overall_success,
            'total_duration': total_duration,
            'test_suites': self.results,
            'summary': {
                'total_tests': total_tests,
                'passed_tests': passed_tests,
                'required_tests': required_tests,
                'passed_required': passed_required
            },
            'quality_metrics': quality_metrics
        }
        
        report_file = project_root / "test_report.json"
        with open(report_file, 'w') as f:
            json.dump(report_data, f, indent=2)
        
        print(f"\n📄 Detailed report saved to: {report_file}")
        
        return overall_success
    
    def print_help(self):
        """Print help information."""
        print("PDF Translation Test Runner")
        print("=" * 40)
        print("\nUsage: python run_tests.py [options]")
        print("\nOptions:")
        print("  -h, --help              Show this help message")
        print("  -a, --all              Run all test suites (default)")
        print("  -q, --quality          Run quality metrics tests")
        print("  -i, --integration      Run integration tests")
        print("  -e, --errors           Run error handling tests")
        print("  -o, --output           Run output generation tests")
        print("  -v, --validation       Run test data validation")
        print("  -l, --list             List available test suites")
        print("  -r, --report           Generate report only")
        print("\nExamples:")
        print("  python run_tests.py                    # Run all tests")
        print("  python run_tests.py --quality          # Run quality tests only")
        print("  python run_tests.py --integration      # Run integration tests only")
        print("  python run_tests.py --list             # List test suites")
        
        print("\nAvailable Test Suites:")
        for name, suite in self.test_suites.items():
            required = " (Required)" if suite['required'] else ""
            print(f"  {name}: {suite['name']}{required}")
    
    def list_suites(self):
        """List all available test suites."""
        print("Available Test Suites:")
        print("-" * 40)
        for name, suite in self.test_suites.items():
            required = " (Required)" if suite['required'] else ""
            print(f"{name}: {suite['name']}{required}")
            print(f"     {suite['description']}")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description='PDF Translation Test Runner')
    parser.add_argument('-a', '--all', action='store_true', help='Run all test suites')
    parser.add_argument('-q', '--quality', action='store_true', help='Run quality metrics tests')
    parser.add_argument('-i', '--integration', action='store_true', help='Run integration tests')
    parser.add_argument('-e', '--errors', action='store_true', help='Run error handling tests')
    parser.add_argument('-o', '--output', action='store_true', help='Run output generation tests')
    parser.add_argument('-v', '--validation', action='store_true', help='Run test data validation')
    parser.add_argument('-l', '--list', action='store_true', help='List available test suites')
    parser.add_argument('-r', '--report', action='store_true', help='Generate report only')
    
    args = parser.parse_args()
    
    runner = TestRunner()
    
    # Handle help
    if len(sys.argv) == 1:
        args.all = True  # Default to running all tests
    
    if args.list:
        runner.list_suites()
        return 0
    
    # Check for existing report
    if args.report:
        report_file = project_root / "test_report.json"
        if report_file.exists():
            with open(report_file, 'r') as f:
                report = json.load(f)
            
            print("Existing Test Report:")
            print("=" * 40)
            print(f"Timestamp: {report['timestamp']}")
            print(f"Overall: {'✅ PASS' if report['overall_success'] else '❌ FAIL'}")
            print(f"Duration: {report['total_duration']:.2f}s")
            print(f"Passed: {report['summary']['passed_tests']}/{report['summary']['total_tests']}")
            return 0 if report['overall_success'] else 1
        else:
            print("No existing report found.")
            return 1
    
    # Run tests based on arguments
    success_count = 0
    total_count = 0
    
    if args.all:
        passed, total = runner.run_all_tests()
        success_count = passed
        total_count = total
    else:
        suites_to_run = []
        if args.quality:
            suites_to_run.append('quality')
        if args.integration:
            suites_to_run.extend(['integration-quality', 'integration-errors', 'integration-output', 'integration-e2e'])
        if args.errors:
            suites_to_run.append('integration-errors')
        if args.output:
            suites_to_run.append('integration-output')
        if args.validation:
            suites_to_run.append('validation')
        
        if not suites_to_run:
            print("No test suites specified. Use --help for options.")
            return 1
        
        for suite_name in suites_to_run:
            if runner.run_test_suite(suite_name):
                success_count += 1
            total_count += 1
    
    # Generate report
    overall_success = runner.generate_report()
    
    # Exit with appropriate code
    return 0 if overall_success else 1


if __name__ == "__main__":
    sys.exit(main())