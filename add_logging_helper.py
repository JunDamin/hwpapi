#!/usr/bin/env python3
"""
Helper script to add logging to functions in notebooks
"""
import json
import re
from pathlib import Path

def add_logging_to_function(cell_source, function_name, module_name=""):
    """Add logging to a function if it doesn't already have it"""
    lines = cell_source.strip().split('\n')
    
    # Check if logging is already present
    has_logger = any('logger = get_logger(' in line or 'self.logger' in line for line in lines)
    if has_logger:
        return cell_source  # Already has logging
    
    # Find the function definition line
    func_def_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith('def ') and function_name in line:
            func_def_idx = i
            break
    
    if func_def_idx is None:
        return cell_source  # No function definition found
    
    # Find the end of the docstring
    docstring_end_idx = func_def_idx + 1
    in_docstring = False
    docstring_quotes = None
    
    for i in range(func_def_idx + 1, len(lines)):
        line = lines[i].strip()
        
        if not in_docstring and ('"""' in line or "'''" in line):
            in_docstring = True
            docstring_quotes = '"""' if '"""' in line else "'''"
            # Check if it's a single-line docstring
            if line.count(docstring_quotes) >= 2:
                docstring_end_idx = i + 1
                break
        elif in_docstring and docstring_quotes in line:
            docstring_end_idx = i + 1
            break
        elif not in_docstring and line and not line.startswith('#'):
            # No docstring, function body starts here
            docstring_end_idx = i
            break
    
    # Create logger line
    logger_name = f"{module_name}.{function_name}" if module_name else function_name
    logger_line = f"    logger = get_logger('{logger_name}')"
    debug_line = f"    logger.debug(f\"Calling {function_name}\")"
    
    # Insert logging lines after docstring
    lines.insert(docstring_end_idx, debug_line)
    lines.insert(docstring_end_idx, logger_line)
    
    return '\n'.join(lines)

def process_notebook(notebook_path):
    """Process a notebook to add logging to functions"""
    with open(notebook_path, 'r', encoding='utf-8') as f:
        notebook = json.load(f)
    
    module_name = Path(notebook_path).stem
    modified = False
    
    for cell in notebook['cells']:
        if cell['cell_type'] == 'code':
            source = ''.join(cell['source'])
            
            # Find function definitions
            func_matches = re.finditer(r'^def\s+(\w+)\s*\(', source, re.MULTILINE)
            for match in func_matches:
                function_name = match.group(1)
                
                # Skip special methods and private functions for now
                if function_name.startswith('_'):
                    continue
                
                new_source = add_logging_to_function(source, function_name, module_name)
                if new_source != source:
                    cell['source'] = new_source.split('\n')
                    # Add newline to each line except the last
                    cell['source'] = [line + '\n' for line in cell['source'][:-1]] + [cell['source'][-1]]
                    modified = True
                    print(f"Added logging to function: {function_name}")
                    break  # Process one function per cell to avoid conflicts
    
    if modified:
        with open(notebook_path, 'w', encoding='utf-8') as f:
            json.dump(notebook, f, indent=2, ensure_ascii=False)
        print(f"Updated {notebook_path}")
    else:
        print(f"No changes needed for {notebook_path}")

if __name__ == "__main__":
    notebooks = [
        "nbs/02_api/00_core.ipynb",
        "nbs/02_api/01_actions.ipynb", 
        "nbs/02_api/02_functions.ipynb",
        "nbs/02_api/03_classes.ipynb",
    ]
    
    for notebook in notebooks:
        if Path(notebook).exists():
            print(f"\nProcessing {notebook}...")
            process_notebook(notebook)
        else:
            print(f"Notebook not found: {notebook}")
