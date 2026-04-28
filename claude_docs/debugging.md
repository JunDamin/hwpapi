# Debugging Tips

## Environment Variables

```bash
# Logging Configuration
HWPAPI_LOG_LEVEL=DEBUG      # Set logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
                            # Default: WARNING (production-friendly, only shows warnings/errors)
                            # Use DEBUG or INFO for development/troubleshooting

# Examples:
# Development - show all logs
export HWPAPI_LOG_LEVEL=DEBUG

# Production - only warnings and errors (default if not set)
export HWPAPI_LOG_LEVEL=WARNING

# Quiet mode - only errors and critical
export HWPAPI_LOG_LEVEL=ERROR
```

**Important**: The default log level is `WARNING`, which means normal users only see warnings, errors, and critical messages. This is intentional to avoid cluttering output in production. Set `HWPAPI_LOG_LEVEL=DEBUG` or `INFO` when you need detailed logging for development or troubleshooting.

## Issue: Import errors

```python
# Check what's exported
python -c "import hwpapi.parametersets; print(dir(hwpapi.parametersets))"

# Check __all__ in generated file
grep "__all__" hwpapi/parametersets.py
```

## Issue: Test failures

```bash
# Run with verbose output
python -m pytest tests/test_hparam.py -vv -s

# Run specific test
python -m pytest tests/test_hparam.py::TestClass::test_method -vv

# Show full traceback
python -m pytest tests/test_hparam.py --tb=long
```

## Issue: Duplicate definitions in code

```bash
# Quick check for duplicate method names in generated file
grep -c "def _format_int_value" hwpapi/parametersets.py  # Should be 1

# Find all duplicate methods in a file
python << 'EOF'
import re
from collections import Counter

with open('hwpapi/parametersets.py', encoding='utf-8') as f:
    content = f.read()

# Find all method definitions
methods = re.findall(r'\n    def (\w+)', content)
method_counts = Counter(methods)

# Show duplicates
duplicates = {name: count for name, count in method_counts.items() if count > 1}
if duplicates:
    print("DUPLICATES FOUND:")
    for name, count in duplicates.items():
        print(f"  {name}: {count} times")
else:
    print("No duplicates found.")
EOF

# After fixing, verify
grep -c "def _format_int_value" hwpapi/parametersets.py  # Should be 1
```
