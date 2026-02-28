#!/usr/bin/env python3
"""C2 Fix: Replace os.makedirs() with os.makedirs(exist_ok=True).

Transforms the TOCTOU pattern:
    if dir_var and not os.path.exists(dir_var):
        os.makedirs(dir_var)

Into the atomic pattern:
    if dir_var:
        os.makedirs(dir_var, exist_ok=True)
"""

import re
import sys
import os

# Two-line pattern: if-check + makedirs on next line
PATTERN = re.compile(
    r'^(\s+)if\s+(\w+)\s+and\s+not\s+os\.path\.exists\(\2\):\s*\n'
    r'\1    os\.makedirs\(\2\)\s*$',
    re.MULTILINE
)


def fix_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    count = len(PATTERN.findall(content))

    def replacer(m):
        indent = m.group(1)
        var = m.group(2)
        return f"{indent}if {var}:\n{indent}    os.makedirs({var}, exist_ok=True)"

    new_content = PATTERN.sub(replacer, content)

    if count > 0:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(new_content)

    return count


if __name__ == '__main__':
    targets = [
        'src/docuflow_mcp/extensions/pdf.py',
        'src/docuflow_mcp/extensions/excel.py',
        'src/docuflow_mcp/extensions/ppt.py',
        'src/docuflow_mcp/extensions/html_to_pptx.py',
        'src/docuflow_mcp/extensions/ocr.py',
    ]

    total = 0
    for target in targets:
        path = os.path.join(os.path.dirname(__file__), '..', target)
        path = os.path.normpath(path)
        if not os.path.exists(path):
            # Try relative to cwd
            path = target
        n = fix_file(path)
        if n:
            print(f"  {target}: {n} fixes")
        total += n

    print(f"Total: {total} os.makedirs() calls fixed")
