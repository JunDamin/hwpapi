#!/usr/bin/env python3
"""
Test script for hwpapi logging system
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'hwpapi'))

def test_logging_import():
    """Test that logging module can be imported"""
    try:
        from hwpapi.logging import get_logger, configure_logging, HwpApiLogger
        print("✓ Successfully imported logging modules")
        return True
    except ImportError as e:
        print(f"✗ Failed to import logging modules: {e}")
        return False

def test_logger_creation():
    """Test logger creation and basic functionality"""
    try:
        from hwpapi.logging import get_logger, configure_logging
        
        # Test basic logger creation
        logger = get_logger('test')
        print("✓ Successfully created logger")
        
        # Test logging at different levels
        logger.info("This is an info message")
        logger.debug("This is a debug message (may not be visible)")
        logger.warning("This is a warning message")
        logger.error("This is an error message")
        
        # Test configuration
        configure_logging(level='DEBUG')
        logger.debug("This debug message should now be visible")
        
        print("✓ Basic logging functionality works")
        return True
    except Exception as e:
        print(f"✗ Logger creation failed: {e}")
        return False

def test_singleton_pattern():
    """Test that HwpApiLogger follows singleton pattern"""
    try:
        from hwpapi.logging import HwpApiLogger
        
        logger1 = HwpApiLogger()
        logger2 = HwpApiLogger()
        
        if logger1 is logger2:
            print("✓ Singleton pattern working correctly")
            return True
        else:
            print("✗ Singleton pattern not working - different instances created")
            return False
    except Exception as e:
        print(f"✗ Singleton test failed: {e}")
        return False

def test_core_module_logging():
    """Test that core module can use logging"""
    try:
        # Test that we can import core with logging
        from hwpapi.core import Engine, App
        print("✓ Core module imports successfully with logging")
        
        # Note: We can't test actual Engine/App creation without HWP installed
        # But we can verify the import works
        return True
    except ImportError as e:
        print(f"✗ Core module import failed: {e}")
        return False

def main():
    """Run all tests"""
    print("Testing hwpapi logging system...")
    print("=" * 50)
    
    tests = [
        test_logging_import,
        test_logger_creation,
        test_singleton_pattern,
        test_core_module_logging,
    ]
    
    results = []
    for test in tests:
        print(f"\nRunning {test.__name__}:")
        result = test()
        results.append(result)
    
    print("\n" + "=" * 50)
    print("Test Summary:")
    passed = sum(results)
    total = len(results)
    print(f"Passed: {passed}/{total}")
    
    if passed == total:
        print("🎉 All tests passed!")
        return 0
    else:
        print("❌ Some tests failed!")
        return 1

if __name__ == "__main__":
    exit(main())
