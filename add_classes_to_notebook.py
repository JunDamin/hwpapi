"""Add generated ParameterSet classes to the notebook."""
import json
from pathlib import Path

def add_classes_to_notebook():
    """Add generated classes to the parameters notebook."""
    notebook_path = Path('nbs/02_api/02_parameters.ipynb')
    generated_file = Path('generated_classes.txt')

    # Read the notebook
    with notebook_path.open('r', encoding='utf-8') as f:
        notebook = json.load(f)

    # Read generated classes (skip the header lines)
    generated_code = generated_file.read_text(encoding='utf-8')
    # Remove the header
    lines = generated_code.split('\n')
    code_lines = []
    started = False
    for line in lines:
        if line.startswith('class '):
            started = True
        if started:
            code_lines.append(line)

    code_content = '\n'.join(code_lines).strip()

    # Prepend export directive
    code_content = '#| export\n\n' + code_content

    # Create a new code cell
    new_cell = {
        'cell_type': 'code',
        'execution_count': None,
        'metadata': {},
        'outputs': [],
        'source': code_content
    }

    # Find the last code cell with classes (should be near the end)
    # Insert before the test cells
    insert_index = len(notebook['cells']) - 1
    for i in range(len(notebook['cells']) - 1, -1, -1):
        cell = notebook['cells'][i]
        if cell['cell_type'] == 'code' and 'class Table(ParameterSet)' in ''.join(cell.get('source', [])):
            insert_index = i + 1
            break

    # Insert the new cell
    notebook['cells'].insert(insert_index, new_cell)

    # Write back to notebook
    with notebook_path.open('w', encoding='utf-8') as f:
        json.dump(notebook, f, ensure_ascii=False, indent=1)

    print(f"Added {len([l for l in code_lines if l.startswith('class ')])} classes to notebook at cell index {insert_index}")
    print(f"Notebook now has {len(notebook['cells'])} cells")

if __name__ == '__main__':
    add_classes_to_notebook()
