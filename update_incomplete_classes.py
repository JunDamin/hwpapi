"""Update incomplete manually created classes with complete versions."""
import json
import re
from pathlib import Path
from typing import Dict, List

# Import the existing generation functions
import sys
sys.path.insert(0, str(Path('.').resolve()))
from analyze_docs import (
    parse_markdown_file,
    generate_class_code,
    clean_property_name,
    clean_description,
    type_to_property
)

CLASSES_TO_UPDATE = [
    'BorderFill',
    'CharShape',
    'FindReplace',
    'ParaShape',
    'BulletShape',
    'Caption',
    'Cell',
    'Table',
    'DrawImageAttr',
    'ShapeObject',
    'NumberingShape'
]

def generate_complete_class(class_name: str, docs_dir: Path) -> str:
    """Generate complete class code from markdown documentation."""
    md_file = docs_dir / f"{class_name}.md"
    if not md_file.exists():
        print(f"Warning: No markdown file found for {class_name}")
        return None

    _, properties = parse_markdown_file(md_file)
    if not properties:
        print(f"Warning: No properties found for {class_name}")
        return None

    return generate_class_code(class_name, properties)

def find_and_replace_class_in_notebook(notebook_path: Path, class_name: str, new_class_code: str) -> bool:
    """Find and replace a class definition in the notebook."""
    with notebook_path.open('r', encoding='utf-8') as f:
        notebook = json.load(f)

    replaced = False
    for cell_idx, cell in enumerate(notebook['cells']):
        if cell.get('cell_type') != 'code':
            continue

        source = ''.join(cell.get('source', []))

        # Check if this cell contains the class
        if f'class {class_name}(ParameterSet):' in source:
            # Find the class definition and replace it
            # Match from class definition to either next class or end of cell
            pattern = rf'(class {class_name}\(ParameterSet\):.*?)(?=\n\nclass \w+\(ParameterSet\):|\Z)'

            if re.search(pattern, source, re.DOTALL):
                # Replace the class
                new_source = re.sub(pattern, new_class_code, source, count=1, flags=re.DOTALL)
                notebook['cells'][cell_idx]['source'] = new_source
                replaced = True
                print(f"[OK] Replaced {class_name} in cell {cell_idx}")
                break

    if replaced:
        # Write back to notebook
        with notebook_path.open('w', encoding='utf-8') as f:
            json.dump(notebook, f, ensure_ascii=False, indent=1)

    return replaced

def main():
    """Main update function."""
    docs_dir = Path('parameterset_docs/markdowns')
    notebook_path = Path('nbs/02_api/02_parameters.ipynb')

    print("="*80)
    print("UPDATING INCOMPLETE CLASSES")
    print("="*80)

    updated_count = 0
    not_found_count = 0

    for class_name in CLASSES_TO_UPDATE:
        print(f"\nProcessing {class_name}...")

        # Generate complete class code
        new_class_code = generate_complete_class(class_name, docs_dir)
        if not new_class_code:
            not_found_count += 1
            continue

        # Replace in notebook
        if find_and_replace_class_in_notebook(notebook_path, class_name, new_class_code):
            updated_count += 1
        else:
            print(f"  [FAIL] Could not find {class_name} in notebook")
            not_found_count += 1

    print("\n" + "="*80)
    print(f"SUMMARY: Updated {updated_count}/{len(CLASSES_TO_UPDATE)} classes")
    if not_found_count:
        print(f"  Could not update {not_found_count} classes")
    print("="*80)

if __name__ == '__main__':
    main()
