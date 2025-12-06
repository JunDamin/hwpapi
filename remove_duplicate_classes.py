"""Remove duplicate/simplified class definitions from the notebook."""
import json
import re
from pathlib import Path
from collections import defaultdict

CLASSES_WITH_DUPLICATES = [
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
    'DrawFillAttr'
]

def find_class_cells(notebook_path: Path) -> dict:
    """Find all cells containing each class definition."""
    with notebook_path.open('r', encoding='utf-8') as f:
        notebook = json.load(f)

    class_cells = defaultdict(list)

    for idx, cell in enumerate(notebook['cells']):
        if cell.get('cell_type') != 'code':
            continue

        source = ''.join(cell.get('source', []))

        for class_name in CLASSES_WITH_DUPLICATES:
            if f'class {class_name}(ParameterSet):' in source:
                # Count properties to determine if it's the simple or complete version
                prop_count = len(re.findall(r'^\s+\w+\s+=\s+', source, re.MULTILINE))
                class_cells[class_name].append({
                    'cell_idx': idx,
                    'prop_count': prop_count,
                    'source_preview': source[:200]
                })

    return class_cells

def remove_simple_versions(notebook_path: Path, class_cells: dict):
    """Remove the simplified versions (keep the ones with more properties)."""
    with notebook_path.open('r', encoding='utf-8') as f:
        notebook = json.load(f)

    cells_to_remove = set()

    for class_name, occurrences in class_cells.items():
        if len(occurrences) <= 1:
            continue

        # Sort by property count - keep the one with most properties
        occurrences.sort(key=lambda x: x['prop_count'], reverse=True)

        print(f"\n{class_name}:")
        for i, occ in enumerate(occurrences):
            status = "KEEP" if i == 0 else "REMOVE"
            print(f"  Cell {occ['cell_idx']}: {occ['prop_count']} properties [{status}]")

            if i > 0:  # Not the first (most complete) one
                cells_to_remove.add(occ['cell_idx'])

    if not cells_to_remove:
        print("\nNo duplicate cells to remove!")
        return

    # Remove cells (in reverse order to maintain indices)
    print(f"\nRemoving {len(cells_to_remove)} cells...")
    for idx in sorted(cells_to_remove, reverse=True):
        del notebook['cells'][idx]

    # Write back
    with notebook_path.open('w', encoding='utf-8') as f:
        json.dump(notebook, f, ensure_ascii=False, indent=1)

    print(f"[OK] Removed {len(cells_to_remove)} duplicate class definitions")

def main():
    """Main function."""
    notebook_path = Path('nbs/02_api/02_parameters.ipynb')

    print("="*80)
    print("FINDING DUPLICATE CLASS DEFINITIONS")
    print("="*80)

    class_cells = find_class_cells(notebook_path)

    print("\n" + "="*80)
    print("REMOVING SIMPLIFIED VERSIONS")
    print("="*80)

    remove_simple_versions(notebook_path, class_cells)

if __name__ == '__main__':
    main()
