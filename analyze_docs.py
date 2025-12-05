"""Analyze parameterset documentation and generate missing classes."""
import json
import re
from pathlib import Path
from typing import Dict, List, Tuple

def parse_markdown_file(md_path: Path) -> Tuple[str, List[Dict[str, str]]]:
    """Parse a markdown file and extract class name and properties."""
    content = md_path.read_text(encoding='utf-8')

    # Extract class name from first heading
    class_match = re.search(r'^###\s+(\w+)', content, re.MULTILINE)
    if not class_match:
        return None, []

    class_name = class_match.group(1)

    # Extract table rows
    properties = []
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
                properties.append({
                    'name': parts[0],
                    'type': parts[1],
                    'subtype': parts[2],
                    'description': parts[3]
                })
        elif in_table and not line.startswith('|'):
            break

    return class_name, properties

def get_existing_classes_from_notebook(notebook_path: Path) -> set:
    """Extract existing class names from the notebook."""
    with notebook_path.open('r', encoding='utf-8') as f:
        notebook = json.load(f)

    classes = set()
    for cell in notebook.get('cells', []):
        if cell.get('cell_type') == 'code':
            source = ''.join(cell.get('source', []))
            class_matches = re.findall(r'^class (\w+)\(ParameterSet\):', source, re.MULTILINE)
            classes.update(class_matches)
    return classes

def type_to_property(pit_type: str, subtype: str) -> str:
    """Map PIT type to PropertyDescriptor type."""
    type_map = {
        'PIT_BSTR': 'StringProperty',
        'PIT_UI1': 'IntProperty',
        'PIT_UI2': 'IntProperty',
        'PIT_UI4': 'IntProperty',
        'PIT_I1': 'IntProperty',
        'PIT_I2': 'IntProperty',
        'PIT_I4': 'IntProperty',
        'PIT_I': 'IntProperty',
        'PIT_R4': 'PropertyDescriptor',  # Float, use base class
        'PIT_R8': 'PropertyDescriptor',  # Double, use base class
        'PIT_SET': 'TypedProperty',
        'PIT_ARRAY': 'ListProperty',
    }
    return type_map.get(pit_type, 'PropertyDescriptor')

def clean_markdown_links(text: str) -> str:
    """Remove markdown links and keep only the text."""
    # Remove markdown links like [text](url)
    text = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', text)
    # Remove stray markdown link syntax
    text = re.sub(r'\]\([^\)]+\)', '', text)
    return text

def clean_property_name(name: str) -> str:
    """Clean property name from markdown links."""
    name = clean_markdown_links(name)
    # Remove any remaining brackets
    name = name.replace('[', '').replace(']', '')
    return name.strip()

def clean_description(desc: str) -> str:
    """Clean description text."""
    desc = clean_markdown_links(desc)
    # Remove quotes and apostrophes to avoid string issues
    desc = desc.replace('"', "'").replace("'", "")
    # Normalize whitespace
    desc = ' '.join(desc.split())
    # If description is too short or just backslashes/special chars, use generic description
    if len(desc.strip()) < 3 or desc.strip() in ['\\', '\\\\', '']:
        desc = "Parameter property"
    return desc

def generate_class_code(class_name: str, properties: List[Dict[str, str]]) -> str:
    """Generate Python class code for a ParameterSet."""
    lines = [f'class {class_name}(ParameterSet):']
    lines.append(f'    """{class_name} ParameterSet."""')

    for prop in properties:
        prop_name = clean_property_name(prop['name'])
        pit_type = prop['type']
        subtype = prop['subtype']
        description = clean_description(prop['description'])

        property_type = type_to_property(pit_type, subtype)

        if property_type == 'TypedProperty' and subtype:
            # Extract class name from subtype (remove markdown links)
            subtype_clean = clean_property_name(subtype)
            lines.append(f'    {prop_name} = {property_type}("{prop_name}", r"""{description}""", wrap=lambda: {subtype_clean})')
        else:
            lines.append(f'    {prop_name} = {property_type}("{prop_name}", r"""{description}""")')

    return '\n'.join(lines)

def main():
    """Main analysis function."""
    docs_dir = Path('parameterset_docs/markdowns')
    notebook_file = Path('nbs/02_api/02_parameters.ipynb')

    # Get existing classes from notebook
    existing_classes = get_existing_classes_from_notebook(notebook_file)
    print(f"Found {len(existing_classes)} existing classes")
    print(f"Existing: {sorted(existing_classes)}\n")

    # Parse all markdown files
    documented_classes = {}
    for md_file in docs_dir.glob('*.md'):
        class_name, properties = parse_markdown_file(md_file)
        if class_name and properties:
            documented_classes[class_name] = properties

    print(f"Found {len(documented_classes)} documented classes")
    print(f"Documented: {sorted(documented_classes.keys())}\n")

    # Find missing classes
    missing_classes = set(documented_classes.keys()) - existing_classes
    print(f"\nMissing {len(missing_classes)} classes:")
    print(sorted(missing_classes))

    # Generate code for missing classes
    output_file = Path('generated_classes.txt')
    with output_file.open('w', encoding='utf-8') as f:
        f.write("="*80 + "\n")
        f.write("GENERATED CODE FOR MISSING CLASSES\n")
        f.write("="*80 + "\n\n")

        for class_name in sorted(missing_classes):
            properties = documented_classes[class_name]
            code = generate_class_code(class_name, properties)
            f.write(code + "\n\n")

    print(f"\nGenerated code saved to: {output_file.absolute()}")

if __name__ == '__main__':
    main()
