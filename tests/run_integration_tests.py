#!/usr/bin/env python3
"""
Test runner for PDF integration tests.

Provides a simple way to run specific test suites and view results.
"""

import unittest
import sys
import os
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

def run_integration_tests():
    """Run all integration tests."""
    print("Running PDF Integration Tests...")
    print("=" * 50)
    
    # Import the test module
    try:
        from tests.test_pdf_integration import create_integration_test_suite
    except ImportError as e:
        print(f"Error importing test module: {e}")
        return False
    
    # Create and run the test suite
    suite = create_integration_test_suite()
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    # Print summary
    print("\n" + "=" * 50)
    print("INTEGRATION TEST SUMMARY")
    print("=" * 50)
    print(f"Tests run: {result.testsRun}")
    print(f"Failures: {len(result.failures)}")
    print(f"Errors: {len(result.errors)}")
    
    if result.failures:
        print("\nFAILURES:")
        for test, traceback in result.failures:
            print(f"  {test}: {traceback.split('AssertionError: ')[-1].split('\\n')[0]}")
    
    if result.errors:
        print("\nERRORS:")
        for test, traceback in result.errors:
            print(f"  {test}: {traceback.split(': ')[-1].split('\\n')[0]}")
    
    success = result.wasSuccessful()
    print(f"\nResult: {'PASS' if success else 'FAIL'}")
    
    return success

def run_quality_metric_tests():
    """Run only quality metric tests."""
    print("Running Quality Metric Tests...")
    print("=" * 50)
    
    # Import and run specific test class
    from tests.test_pdf_integration import TestQualityMetricsEnforcement
    
    suite = unittest.TestLoader().loadTestsFromTestCase(TestQualityMetricsEnforcement)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    print(f"\nQuality Metric Tests: {'PASS' if result.wasSuccessful() else 'FAIL'}")
    return result.wasSuccessful()

def run_end_to_end_tests():
    """Run only end-to-end tests."""
    print("Running End-to-End Tests...")
    print("=" * 50)
    
    from tests.test_pdf_integration import TestPDFEndToEndTranslation
    
    suite = unittest.TestLoader().loadTestsFromTestCase(TestPDFEndToEndTranslation)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    print(f"\nEnd-to-End Tests: {'PASS' if result.wasSuccessful() else 'FAIL'}")
    return result.wasSuccessful()

def main():
    """Main entry point."""
    if len(sys.argv) > 1:
        test_type = sys.argv[1].lower()
        
        if test_type == "quality":
            success = run_quality_metric_tests()
        elif test_type == "e2e":
            success = run_end_to_end_tests()
        elif test_type == "integration":
            success = run_integration_tests()
        else:
            print(f"Unknown test type: {test_type}")
            print("Available types: quality, e2e, integration")
            sys.exit(1)
    else:
        # Default: run all integration tests
        success = run_integration_tests()
    
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()