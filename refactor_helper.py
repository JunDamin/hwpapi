#!/usr/bin/env python3
"""
Helper script to refactor parameter set classes to use new property descriptors.
"""

import re

# Mapping of old static methods to new property classes
METHOD_MAPPING = {
    'ParameterSet._int_prop': 'IntProperty',
    'ParameterSet._bool_prop': 'BoolProperty', 
    'ParameterSet._str_prop': 'StringProperty',
    'ParameterSet._color_prop': 'ColorProperty',
    'ParameterSet._unit_prop': 'UnitProperty',
    'ParameterSet._mapped_prop': 'MappedProperty',
    'ParameterSet._typed_prop': 'TypedProperty',
    'ParameterSet._int_list_prop': 'ListProperty',
    'ParameterSet._tuple_list_prop': 'ListProperty',
    'ParameterSet._gradation_color_prop': 'ListProperty',
}

def refactor_property_definition(line):
    """Refactor a single property definition line."""
    for old_method, new_class in METHOD_MAPPING.items():
        if old_method in line:
            # Extract property name and arguments
            prop_match = re.match(r'(\s*)(\w+)\s*=\s*ParameterSet\._\w+_prop\((.+)\)', line)
            if prop_match:
                indent, prop_name, args = prop_match.groups()
                
                # Handle special cases
                if old_method == 'ParameterSet._unit_prop':
                    # UnitProperty has different argument order: key, unit, doc
                    args_parts = [arg.strip() for arg in args.split(',', 2)]
                    if len(args_parts) >= 3:
                        key, unit_or_doc, doc_or_unit = args_parts
                        if 'unit=' in args:
                            # Format: key, doc, unit=unit
                            args = f"{key}, {unit_or_doc.split('unit=')[1].strip()}, {doc_or_unit}"
                        else:
                            # Format: key, unit, doc
                            args = f"{key}, {unit_or_doc}, {doc_or_unit}"
                
                elif old_method == 'ParameterSet._mapped_prop':
                    # MappedProperty: key, mapping, doc
                    # Sometimes has mapping= parameter
                    if 'mapping=' in args:
                        args = args.replace('mapping=', '')
                    # Sometimes has doc= parameter
                    if 'doc=' in args:
                        args = args.replace('doc=', '')
                
                elif old_method in ['ParameterSet._int_list_prop', 'ParameterSet._tuple_list_prop']:
                    # These become ListProperty with item_type
                    args_parts = [arg.strip() for arg in args.split(',', 1)]
                    if len(args_parts) >= 2:
                        key, doc = args_parts
                        item_type = 'int' if 'int_list' in old_method else 'tuple'
                        args = f"{key}, {doc}, item_type={item_type}"
                
                elif old_method == 'ParameterSet._gradation_color_prop':
                    # Becomes ListProperty with min/max length
                    args_parts = [arg.strip() for arg in args.split(',', 1)]
                    if len(args_parts) >= 2:
                        key, doc = args_parts
                        args = f"{key}, {doc}, min_length=2, max_length=10"
                
                return f"{indent}{prop_name} = {new_class}({args})"
    
    return line

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
            if line.strip().startswith('def ') or (line.strip() and not line.startswith('    ')):
                skip_init = False
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
    print("This is a helper script for refactoring parameter sets.")
    print("Use the functions in this script to help with the refactoring process.")
