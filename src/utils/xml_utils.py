"""XML parsing utilities"""
import xml.etree.ElementTree as ET
from typing import Optional, Dict, List, Tuple
from ..constants import XMLNamespaces

def find_elements(root: ET.Element, path: str, namespace: str = XMLNamespaces.MAIN) -> List[ET.Element]:
    """Find all elements matching the path with namespace"""
    ns = {'main': namespace}
    return root.findall(path, ns)

def get_attribute(element: ET.Element, attr: str, default: str = '') -> str:
    """Safely get attribute value"""
    return element.get(attr, default)

def parse_cell_reference(cell_ref: str) -> Tuple[str, int]:
    """Parse cell reference into column and row"""
    col = ''.join(c for c in cell_ref if c.isalpha())
    row = int(''.join(c for c in cell_ref if c.isdigit()))
    return col, row 