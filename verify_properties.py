"""Verify that ParameterSet property names match COM object attributes."""
import re
import json
from pathlib import Path
from typing import Dict, Set, List
from collections import defaultdict

def extract_properties_from_markdown(md_path: Path) -> tuple[str, List[str]]:
    """Extract class name and property Item IDs from markdown."""
    content = md_path.read_text(encoding='utf-8')

    # Extract class name
    class_match = re.search(r'^###\s+(\w+)', content, re.MULTILINE)
    if not class_match:
        return None, []

    class_name = class_match.group(1)

    # Extract Item IDs from table
    item_ids = []
    in_table = False
    for line in content.split('\n'):
        if '| Item ID | Type | SubType | Description |' in line:
            in_table = True
            continue
        if in_table and line.startswith('| ---'):
            continue
        if in_table and line.startswith('|'):
            parts = [p.strip() for p in line.split('|')[1:-1]]
            if len(parts) >= 4 and parts[0] and not parts[0].startswith('---'):
                # Clean markdown links from Item ID
                item_id = parts[0]
                item_id = re.sub(r'\[([^\]]+)\].*', r'\1', item_id)
                item_id = item_id.replace('[', '').replace(']', '')
                item_ids.append(item_id.strip())
        elif in_table and not line.startswith('|'):
            break

    return class_name, item_ids

def extract_properties_from_class(class_source: str) -> List[str]:
    """Extract property keys from ParameterSet class source code."""
    properties = []

    # Match property definitions - handle both old and new style
    # Old: property_name = ParameterSet._int_prop("KeyName", ...)
    # New: PropertyName = IntProperty("KeyName", ...)
    patterns = [
        r'^\s+(\w+)\s+=\s+(?:Int|String|Bool|Typed|Property|Color|Unit|Mapped|List)Property\(["\']([^"\']+)["\']',
        r'^\s+(\w+)\s+=\s+ParameterSet\._\w+_prop\(["\']([^"\']+)["\']',
        r'^\s+(\w+)\s+=\s+PropertyDescriptor\(["\']([^"\']+)["\']',
    ]

    for line in class_source.split('\n'):
        for pattern in patterns:
            match = re.search(pattern, line)
            if match:
                var_name, key_name = match.groups()
                properties.append(key_name)
                break

    return properties

def get_classes_from_notebook(notebook_path: Path) -> Dict[str, List[str]]:
    """Extract all ParameterSet classes and their properties from notebook."""
    with notebook_path.open('r', encoding='utf-8') as f:
        notebook = json.load(f)

    classes = {}
    for cell in notebook.get('cells', []):
        if cell.get('cell_type') == 'code':
            source = ''.join(cell.get('source', []))

            # Find class definitions
            class_matches = re.finditer(r'^class (\w+)\(ParameterSet\):(.+?)(?=^class |\Z)',
                                       source, re.MULTILINE | re.DOTALL)

            for match in class_matches:
                class_name = match.group(1)
                class_body = match.group(2)
                properties = extract_properties_from_class(class_body)
                if properties:
                    classes[class_name] = properties

    return classes

def main():
    """Main verification function."""
    docs_dir = Path('parameterset_docs/markdowns')
    notebook_path = Path('nbs/02_api/02_parameters.ipynb')

    # Get documented properties from markdown
    print("Reading markdown documentation...")
    doc_properties = {}
    for md_file in docs_dir.glob('*.md'):
        class_name, item_ids = extract_properties_from_markdown(md_file)
        if class_name and item_ids:
            doc_properties[class_name] = set(item_ids)

    print(f"Found {len(doc_properties)} documented classes\n")

    # Get implemented properties from notebook
    print("Reading notebook classes...")
    impl_properties = get_classes_from_notebook(notebook_path)
    print(f"Found {len(impl_properties)} implemented classes\n")

    # Compare
    print("="*80)
    print("VERIFICATION RESULTS")
    print("="*80)

    mismatches = []
    missing_props = defaultdict(list)
    extra_props = defaultdict(list)

    for class_name in sorted(set(doc_properties.keys()) & set(impl_properties.keys())):
        doc_props = doc_properties[class_name]
        impl_props = set(impl_properties[class_name])

        missing = doc_props - impl_props
        extra = impl_props - doc_props

        if missing or extra:
            mismatches.append(class_name)
            if missing:
                missing_props[class_name] = sorted(missing)
            if extra:
                extra_props[class_name] = sorted(extra)

    if not mismatches:
        print("[OK] All properties match perfectly!")
        print(f"  Verified {len(set(doc_properties.keys()) & set(impl_properties.keys()))} classes")
    else:
        print(f"[MISMATCH] Found mismatches in {len(mismatches)} classes:\n")

        for class_name in sorted(mismatches):
            print(f"\n{class_name}:")
            if class_name in missing_props:
                print(f"  Missing properties (in docs but not in code):")
                for prop in missing_props[class_name]:
                    print(f"    - {prop}")
            if class_name in extra_props:
                print(f"  Extra properties (in code but not in docs):")
                for prop in extra_props[class_name]:
                    print(f"    - {prop}")

    # Check for classes in docs but not implemented
    not_implemented = set(doc_properties.keys()) - set(impl_properties.keys())
    if not_implemented:
        print(f"\n\n[WARNING] Classes in docs but not implemented: {len(not_implemented)}")
        for name in sorted(not_implemented):
            print(f"  - {name}")

    # Check for classes implemented but not in docs
    not_documented = set(impl_properties.keys()) - set(doc_properties.keys())
    if not_documented:
        print(f"\n\n[WARNING] Classes implemented but not in docs: {len(not_documented)}")
        for name in sorted(not_documented):
            print(f"  - {name}")

if __name__ == '__main__':
    main()
