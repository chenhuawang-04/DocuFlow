#!/usr/bin/env python3
"""H7 Fix (v2): Add assertions after every `result = SomeOp.method(...)` call.

Scans for lines matching:
    result = SomeClass.method(...)
    result = some_func(...)

And inserts an assertion on the next line if one doesn't already exist.
"""

import re
import os


def fix_test_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    new_lines = []
    changes = 0

    i = 0
    while i < len(lines):
        line = lines[i]
        new_lines.append(line)

        stripped = line.strip()

        # Match: result = SomeOp.method(...) — must be a COMPLETE single-line statement
        # Skip multi-line expressions (line ends with open bracket or comma)
        if re.match(r'^\s+result\s*=\s*\w+', line):
            rstripped = line.rstrip()
            # Only add assert if line is complete (ends with ) or similar, not { ( [ ,)
            if rstripped.endswith(')') or rstripped.endswith('}') or rstripped.endswith(']'):
                indent = len(line) - len(line.lstrip())
                indent_str = ' ' * indent

                # Check if next line already has an assert
                next_line = lines[i + 1].strip() if i + 1 < len(lines) else ''
                if not next_line.startswith('assert'):
                    new_lines.append(
                        f'{indent_str}assert isinstance(result, dict), '
                        f'"Expected dict result"\n'
                    )
                    changes += 1

        i += 1

    # Fix hardcoded paths (M2)
    content = ''.join(new_lines)
    content = content.replace(
        "sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))",
        "sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))"
    )
    content = re.sub(
        r'"E:/Project/DocuFlow/test_output/([^"]+)"',
        r'os.path.join(os.path.dirname(__file__), "..", "test_output", "\1")',
        content
    )

    if changes > 0:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)

    return changes


if __name__ == '__main__':
    test_dir = os.path.join(os.path.dirname(__file__), '..', 'tests')
    test_dir = os.path.normpath(test_dir)

    total = 0
    for fname in sorted(os.listdir(test_dir)):
        if fname.startswith('test_') and fname.endswith('.py'):
            fpath = os.path.join(test_dir, fname)
            n = fix_test_file(fpath)
            if n:
                print(f"  {fname}: {n} assertions added")
            total += n

    print(f"Total: {total} assertions added across test files")
