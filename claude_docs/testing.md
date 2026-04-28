# Testing Strategy

## Test Structure

```
tests/test_hparam.py
├── Unit tests (with mocks) - Run without HWP
├── Integration tests - Require HWP installed
└── Graceful skipping - Tests skip if dependencies unavailable
```

## Running Tests

```bash
# All tests
python -m pytest tests/test_hparam.py -v

# Specific test class
python -m pytest tests/test_hparam.py::TestParameterSetUpdateFrom -v

# Show skipped tests
python -m pytest tests/test_hparam.py -v -ra
```

## Test Requirements

**For Unit Tests:** Just Python + pytest
**For Integration Tests:**
- Windows OS
- HWP installed
- pywin32

## Writing New Tests

```python
import unittest
from hwpapi.parametersets import ParameterSet, IntProperty

class TestMyFeature(unittest.TestCase):
    def test_feature(self):
        # Use real ParameterSet subclasses that have properties defined
        from hwpapi.parametersets import CharShape

        ps = CharShape()
        ps.bold = True
        self.assertEqual(ps.bold, True)
```

**Important:** When testing ParameterSet, use classes with actual property descriptors, not manual attributes_names lists.

## Auto-Yes Dialog Mode

Pytest runs must auto-answer Yes to all HWP dialogs (`SetMessageBoxMode 0x111111`) so HWP dialogs never block a test run. This is captured in project memory.

## Quick Reference Commands

```bash
# Run all tests
python -m pytest tests/ -v

# Test specific module
python -m pytest tests/test_hparam.py -v

# Verify imports
python -c "import hwpapi; print('OK')"

# Install in dev mode
pip install -e .
```
