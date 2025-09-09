#!/usr/bin/env python3
"""
Complete the refactoring of parametersets.py by replacing all remaining static method calls
with new property descriptor instantiations.
"""

import re

def refactor_file(file_path):
    """Refactor the entire parametersets.py file."""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Dictionary mapping old static methods to new classes
    replacements = [
        (r'ParameterSet\._int_prop\(([^,]+),\s*([^,]+)(?:,\s*([^,]+))?(?:,\s*([^)]+))?\)', 
         r'IntProperty(\1, \2\3\4)'),
        (r'ParameterSet\._bool_prop\(([^,]+),\s*([^)]+)\)', 
         r'BoolProperty(\1, \2)'),
        (r'ParameterSet\._str_prop\(([^,]+),\s*([^)]+)\)', 
         r'StringProperty(\1, \2)'),
        (r'ParameterSet\._color_prop\(([^,]+),\s*([^)]+)\)', 
         r'ColorProperty(\1, \2)'),
        (r'ParameterSet\._unit_prop\(([^,]+),\s*unit=([^,]+),\s*doc=([^)]+)\)', 
         r'UnitProperty(\1, \2, \3)'),
        (r'ParameterSet\._unit_prop\(([^,]+),\s*([^,]+),\s*([^)]+)\)', 
         r'UnitProperty(\1, \2, \3)'),
        (r'ParameterSet\._mapped_prop\(([^,]+),\s*([^,]+),\s*doc=([^)]+)\)', 
         r'MappedProperty(\1, \2, \3)'),
        (r'ParameterSet\._mapped_prop\(([^,]+),\s*mapping=([^,]+),\s*doc=([^)]+)\)', 
         r'MappedProperty(\1, \2, \3)'),
        (r'ParameterSet\._mapped_prop\(([^,]+),\s*([^,]+),\s*([^)]+)\)', 
         r'MappedProperty(\1, \2, \3)'),
        (r'ParameterSet\._typed_prop\(([^,]+),\s*([^,]+),\s*([^)]+)\)', 
         r'TypedProperty(\1, \2, \3)'),
        (r'ParameterSet\._int_list_prop\(([^,]+),\s*([^)]+)\)', 
         r'ListProperty(\1, \2, item_type=int)'),
        (r'ParameterSet\._tuple_list_prop\(([^,]+),\s*([^)]+)\)', 
         r'ListProperty(\1, \2, item_type=tuple)'),
        (r'ParameterSet\._gradation_color_prop\(([^,]+),\s*([^)]+)\)', 
         r'ListProperty(\1, \2, min_length=2, max_length=10)'),
    ]
    
    # Apply all replacements
    for pattern, replacement in replacements:
        content = re.sub(pattern, replacement, content, flags=re.MULTILINE | re.DOTALL)
    
    # Remove __init__ methods and attributes_names assignments
    content = remove_init_and_attributes(content)
    
    # Write the refactored content back
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"Refactoring completed for {file_path}")

def remove_init_and_attributes(content):
    """Remove __init__ methods and attributes_names definitions."""
    lines = content.split('\n')
    result_lines = []
    skip_init = False
    skip_attributes = False
    brace_count = 0
    
    for line in lines:
        # Skip __init__ method
        if 'def __init__(self, parameterset):' in line:
            skip_init = True
            continue
        elif skip_init:
            if line.strip() and not line.startswith('    '):
                # Check if this is a property definition or another method/class
                if any(prop_class in line for prop_class in ['Property(', 'ParameterSet._']):
                    skip_init = False
                elif line.strip().startswith('def ') or line.strip().startswith('class '):
                    skip_init = False
                elif line.strip().startswith('#'):
                    skip_init = False
                else:
                    continue
            elif line.strip().startswith('def ') or line.strip().startswith('class '):
                skip_init = False
            elif not line.strip():
                continue
            else:
                continue
        
        # Skip attributes_names assignment
        if 'self.attributes_names = [' in line:
            skip_attributes = True
            brace_count = line.count('[') - line.count(']')
            continue
        elif skip_attributes:
            brace_count += line.count('[') - line.count(']')
            if brace_count <= 0:
                skip_attributes = False
            continue
        
        if not skip_init and not skip_attributes:
            result_lines.append(line)
    
    return '\n'.join(result_lines)

if __name__ == "__main__":
    refactor_file('hwpapi/parametersets.py')
